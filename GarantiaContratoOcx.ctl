VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl GarantiaContratoOcx 
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   ScaleHeight     =   4425
   ScaleWidth      =   7170
   Begin VB.ComboBox ServicoPeca 
      Height          =   315
      Left            =   615
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2295
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Contrato 
      Height          =   225
      Left            =   5325
      TabIndex        =   4
      Top             =   2355
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   10
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Garantia 
      Height          =   225
      Left            =   3960
      TabIndex        =   5
      Top             =   2340
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   225
      Left            =   2385
      TabIndex        =   3
      Top             =   2355
      Width           =   1500
      _ExtentX        =   2646
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
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2160
      Picture         =   "GarantiaContratoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3780
      Width           =   1005
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   3570
      Picture         =   "GarantiaContratoOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3780
      Width           =   1005
   End
   Begin MSFlexGridLib.MSFlexGrid GridGarantiaContrato 
      Height          =   3450
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6085
      _Version        =   393216
      Rows            =   6
      Cols            =   3
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
   End
End
Attribute VB_Name = "GarantiaContratoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iTipoAlterado As Integer
Dim iAlterado As Integer


Dim gcolGarantiaContratoSRVOrigem As Collection
Dim gcolGarantiaContratoSRVDestino As Collection
Dim objGridGarantiaContrato As AdmGrid

Dim iGrid_ServicoPeca_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Garantia_Col As Integer
Dim iGrid_Contrato_Col As Integer

Dim giFrameAtual As Integer

'Public Function Trata_Parametros(ByVal colServicoPecaOrigem As Collection, colGarantiaContratoSRV As Collection) As Long
''Trata os parametros passados para a tela..
'
'Dim lErro As Long
'Dim sProduto As String
'Dim objGarantiaContratoSRV As ClassGarantiaContratoSRV
'Dim iIndice As Integer
'Dim vServicoPeca As Variant
'Dim sServicoPeca As String
'
'
'On Error GoTo Erro_Trata_Parametros
'
'    Set gcolGarantiaContratoSRVDestino = colGarantiaContratoSRV
'
'    For Each vServicoPeca In colServicoPecaOrigem
'
'        lErro = Mascara_RetornaProdutoTela(vServicoPeca, sServicoPeca)
'        If lErro <> SUCESSO Then gError 188079
'
'        ServicoPeca.AddItem sServicoPeca
'
'    Next
'
'    For iIndice = 1 To colGarantiaContratoSRV.Count
'
'        Set objGarantiaContratoSRV = colGarantiaContratoSRV(iIndice)
'
'        lErro = Mascara_RetornaProdutoTela(objGarantiaContratoSRV.sServicoPecaSRV, sServicoPeca)
'        If lErro <> SUCESSO Then gError 188080
'
'        GridGarantiaContrato.TextMatrix(iIndice, iGrid_ServicoPeca_Col) = sServicoPeca
'        GridGarantiaContrato.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objGarantiaContratoSRV.dQuantidade)
'        GridGarantiaContrato.TextMatrix(iIndice, iGrid_Garantia_Col) = CStr(objGarantiaContratoSRV.lGarantiaCod)
'        GridGarantiaContrato.TextMatrix(iIndice, iGrid_Contrato_Col) = objGarantiaContratoSRV.sContratoCod
'
'    Next
'
'    'Atualiza o número de linhas existentes
'    objGridGarantiaContrato.iLinhasExistentes = colGarantiaContratoSRV.Count
'
'    iAlterado = 0
'
'    Trata_Parametros = SUCESSO
'
'    Exit Function
'
'Erro_Trata_Parametros:
'
'    Trata_Parametros = gErr
'
'    Select Case gErr
'
'        Case 188079 To 188080
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188081)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Form_Load()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Form_Load
'
'
'    Set objGridGarantiaContrato = New AdmGrid
'
'    Call Inicializa_Grid_GarantiaContrato(objGridGarantiaContrato)
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Function
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188082)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Inicializa_Grid_GarantiaContrato(objGridInt As AdmGrid) As Long
''Inicializa o Grid
'
'    'Form do Grid
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add ("Item")
'    objGridInt.colColuna.Add ("Serviço/Peça")
'    objGridInt.colColuna.Add ("Quantidade")
'    objGridInt.colColuna.Add ("Garantia")
'    objGridInt.colColuna.Add ("Contrato")
'
'    objGridInt.colCampo.Add (ServicoPeca.Name)
'    objGridInt.colCampo.Add (Quantidade.Name)
'    objGridInt.colCampo.Add (Garantia.Name)
'    objGridInt.colCampo.Add (Contrato.Name)
'
'    'Controles que participam do Grid
'    iGrid_ServicoPeca_Col = 1
'    iGrid_Quantidade_Col = 2
'    iGrid_Garantia_Col = 3
'    iGrid_Contrato_Col = 4
'
'    objGridInt.objGrid = GridGarantiaContrato
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAXIMO_GARANTIA_CONTRATO + 1
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 9
'
'    'Largura da primeira coluna
'    GridGarantiaContrato.ColWidth(0) = 400
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_GarantiaContrato = SUCESSO
'
'End Function
'
'Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
'End Sub
'
'Public Sub Form_Unload(Cancel As Integer)
'    Call ComandoSeta_Liberar(Me.Name)
'End Sub
'
'Private Sub BotaoCancela_Click()
'    Unload Me
'End Sub
'
'Public Sub Form_Activate()
''    Call TelaIndice_Preenche(Me)
'End Sub
'
''***************************************************
''Trecho de codigo comum as telas
''***************************************************
'
'Public Function Form_Load_Ocx() As Object
''    ??? Parent.HelpContextID = IDH_
'    Set Form_Load_Ocx = Me
'    Caption = "Garantia/Contrato"
'    Call Form_Load
'End Function
'
'Public Function Name() As String
'    Name = "GarantiaContrato"
'End Function
'
'Public Sub Show()
''    Parent.Show
''    Parent.SetFocus
'End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Controls
'Public Property Get Controls() As Object
'    Set Controls = UserControl.Controls
'End Property
'
'Public Property Get hWnd() As Long
'    hWnd = UserControl.hWnd
'End Property
'
'Public Property Get Height() As Long
'    Height = UserControl.Height
'End Property
'
'Public Property Get Width() As Long
'    Width = UserControl.Width
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ActiveControl
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = UserControl.ActiveControl
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
''''    m_Caption = New_Caption
'End Property
'
''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'End Sub
'
'Private Sub Unload(objme As Object)
'    RaiseEvent Unload
'End Sub
'
'Private Sub GridGarantiaContrato_Click()
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGridGarantiaContrato, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridGarantiaContrato, iAlterado)
'
'    End If
'
'End Sub
'
'Private Sub GridGarantiaContrato_EnterCell()
'
'    Call Grid_Entrada_Celula(objGridGarantiaContrato, iAlterado)
'
'End Sub
'
'Private Sub GridGarantiaContrato_GotFocus()
'
'    Call Grid_Recebe_Foco(objGridGarantiaContrato)
'
'End Sub
'
'Private Sub GridGarantiaContrato_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridGarantiaContrato, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridGarantiaContrato, iAlterado)
'    End If
'
'End Sub
'
'Private Sub GridGarantiaContrato_LeaveCell()
'
'    Call Saida_Celula(objGridGarantiaContrato)
'
'End Sub
'
'Private Sub GridGarantiaContrato_Validate(Cancel As Boolean)
'
'    Call Grid_Libera_Foco(objGridGarantiaContrato)
'
'End Sub
'
'Private Sub GridGarantiaContrato_Scroll()
'
'    Call Grid_Scroll(objGridGarantiaContrato)
'
'End Sub
'
'Private Sub GridGarantiaContrato_RowColChange()
'
'    Call Grid_RowColChange(objGridGarantiaContrato)
'
'End Sub
'
'Private Sub GridGarantiaContrato_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Call Grid_Trata_Tecla1(KeyCode, objGridGarantiaContrato)
'
'End Sub
'
'Public Sub ServicoPeca_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ServicoPeca_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridGarantiaContrato)
'
'End Sub
'
'Public Sub ServicoPeca_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGarantiaContrato)
'
'End Sub
'
'Public Sub ServicoPeca_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridGarantiaContrato.objControle = ServicoPeca
'    lErro = Grid_Campo_Libera_Foco(objGridGarantiaContrato)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Quantidade_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Quantidade_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridGarantiaContrato)
'
'End Sub
'
'Public Sub Quantidade_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGarantiaContrato)
'
'End Sub
'
'Public Sub Quantidade_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridServico.objControle = Quantidade
'    lErro = Grid_Campo_Libera_Foco(objGridGarantiaContrato)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Garantia_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Garantia_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridGarantiaContrato)
'
'End Sub
'
'Public Sub Garantia_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGarantiaContrato)
'
'End Sub
'
'Public Sub Garantia_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridGarantiaContrato.objControle = Garantia
'    lErro = Grid_Campo_Libera_Foco(objGridGarantiaContrato)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Contrato_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Contrato_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridGarantiaContrato)
'
'End Sub
'
'Public Sub Contrato_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGarantiaContrato)
'
'End Sub
'
'Public Sub Contrato_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridGarantiaContrato.objControle = Contrato
'    lErro = Grid_Campo_Libera_Foco(objGridGarantiaContrato)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Function Saida_Celula(objGridInt As AdmGrid) As Long
''Faz a critica da célula do grid que está deixando de ser a corrente
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula
'
'    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
'
'    If lErro = SUCESSO Then
'
'        'Verifica qual a coluna atual do Grid
'        Select Case objGridInt.objGrid.Col
'
'            'Se for a de ServicoPeca
'            Case iGrid_ServicoPeca_Col
'                lErro = Saida_Celula_ServicoPeca(objGridInt)
'                If lErro <> SUCESSO Then gError 188083
'
'            'Se for a de Quantidade
'            Case iGrid_Quantidade_Col
'                lErro = Saida_Celula_Quantidade(objGridInt)
'                If lErro <> SUCESSO Then gError 188084
'
'            'Se for a de Garantia
'            Case iGrid_Garantia_Col
'                lErro = Saida_Celula_Garantia(objGridInt)
'                If lErro <> SUCESSO Then gError 188085
'
'            'Se for a de Contrato
'            Case iGrid_Contrato_Col
'                lErro = Saida_Celula_Contrato(objGridInt)
'                If lErro <> SUCESSO Then gError 188086
'
'        End Select
'
'
'    End If
'
'
'    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 188087
'
'    Saida_Celula = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula:
'
'    Saida_Celula = gErr
'
'    Select Case gErr
'
'        Case 188083 To 188017
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188088)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_ServicoPeca(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_ServicoPeca
'
'    Set objGridInt.objControle = ServicoPeca
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 188089
'
'    If Len(Trim(ServicoPeca.Text)) <> 0 Then
'
'        If GridGarantiaContrato.Row - GridGarantiaContrato.FixedRows = objGridGarantiaContrato.iLinhasExistentes Then
'
'            objGridGarantiaContrato.iLinhasExistentes = objGridGarantiaContrato.iLinhasExistentes + 1
'
'        End If
'
'    End If
'
'    Saida_Celula_ServicoPeca = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_ServicoPeca:
'
'    Saida_Celula_ServicoPeca = gErr
'
'    Select Case gErr
'
'        Case 188089
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188090)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Quantidadeque está deixando de ser a corrente
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_Quantidade
'
'    Set objGridInt.objControle = Quantidade
'
'    If Len(Quantidade.Text) > 0 Then
'
'        lErro = Valor_Positivo_Critica(Quantidade.Text)
'        If lErro <> SUCESSO Then gError 188091
'
'        Quantidade.Text = Formata_Estoque(Quantidade.Text)
'
'    End If
'
'    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 188092
'
'    If Len(Quantidade.Text) > 0 Then
'
'        If GridGarantiaContrato.Row - GridGarantiaContrato.FixedRows = objGridGarantiaContrato.iLinhasExistentes Then
'
'            objGridGarantiaContrato.iLinhasExistentes = objGridGarantiaContrato.iLinhasExistentes + 1
'
'        End If
'
'    End If
'
'    Saida_Celula_Quantidade = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Quantidade:
'
'    Saida_Celula_Quantidade = gErr
'
'    Select Case gErr
'
'        Case 188091, 188092
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188093)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_Garantia(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'Dim iProdutoPreenchido As Integer
'Dim objProduto As New ClassProduto
'Dim vbMsg As VbMsgBoxResult
'
'On Error GoTo Erro_Saida_Celula_Garantia
'
'    Set objGridInt.objControle = Garantia
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 188024
'
'    If Len(Trim(Garantia.Text)) > 0 Then
'
'        If GridGarantiaContrato.Row - GridGarantiaContrato.FixedRows = objGridGarantiaContrato.iLinhasExistentes Then
'
'            objGridGarantiaContrato.iLinhasExistentes = objGridGarantiaContrato.iLinhasExistentes + 1
'
'        End If
'
'    End If
'
'    Saida_Celula_Garantia = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Garantia:
'
'    Saida_Celula_Garantia = gErr
'
'    Select Case gErr
'
'        Case 188024
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188025)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub BotaoOK_Click()
'
'Dim lErro As Long
'Dim objGarantia As New ClassGarantia
'
'On Error GoTo Erro_BotaoOK_Click
'
'    lErro = Valida_Dados_Tela()
'    If lErro <> SUCESSO Then gError 188026
'
'    'Move os dados da tela para o objRelacionamentoClie
'    lErro = Move_ProdutoSRV_Memoria()
'    If lErro <> SUCESSO Then gError 188027
'
'    Unload Me
'
'    Exit Sub
'
'Erro_BotaoOK_Click:
'
'    Select Case gErr
'
'        Case 188026, 188027
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188028)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function Valida_Dados_Tela() As Long
''Verifica se os dados da tela são válidos
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim dQuantidade As Double
'Dim objProdutoSRV As ClassProdutoSRV
'Dim colProdutoSRV As New Collection
'Dim iAchou As Integer
'Dim objProdutoSRVOrigem As ClassProdutoSRV
'Dim sServicoFormatado As String
'Dim iServicoPreenchido As Integer
'Dim sProdutoSRV As String
'Dim iProdutoPreenchido As Integer
'Dim iIndice1 As Integer
'Dim sServicoFormatado1 As String
'Dim iServicoPreenchido1 As Integer
'Dim sProdutoSRV1 As String
'Dim iProdutoPreenchido1 As Integer
'
'On Error GoTo Erro_Valida_Dados_Tela
'
'    For iIndice = 1 To objGridServico.iLinhasExistentes
'
'        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Servico_Col))) = 0 Then gError 188029
'
'        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 188030
'
'        If StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col)) <= 0 Then gError 188031
'
'        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Then gError 188032
'
'        iAchou = 0
'
'        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Servico_Col), sServicoFormatado, iServicoPreenchido)
'        If lErro <> SUCESSO Then gError 188033
'
'        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoSRV, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 188034
'
'
'        For iIndice1 = iIndice + 1 To objGridServico.iLinhasExistentes
'
'            lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice1, iGrid_Servico_Col), sServicoFormatado1, iServicoPreenchido1)
'            If lErro <> SUCESSO Then gError 188035
'
'            lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice1, iGrid_Produto_Col), sProdutoSRV1, iProdutoPreenchido1)
'            If lErro <> SUCESSO Then gError 188036
'
'            If sServicoFormatado1 = sServicoFormatado And sProdutoSRV1 = sProdutoSRV Then gError 188037
'
'        Next
'
'        For Each objProdutoSRV In colProdutoSRV
'            If objProdutoSRV.sProdutoSRV = sProdutoSRV Then
'                objProdutoSRV.dQuantidade = objProdutoSRV.dQuantidade + StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))
'                iAchou = 1
'                Exit For
'            End If
'        Next
'
'        If iAchou = 0 Then
'
'            Set objProdutoSRV = New ClassProdutoSRV
'
'            objProdutoSRV.sProdutoSRV = sProdutoSRV
'            objProdutoSRV.dQuantidade = StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))
'
'            colProdutoSRV.Add objProdutoSRV
'
'        End If
'
'    Next
'
'    For Each objProdutoSRV In colProdutoSRV
'
'        For Each objProdutoSRVOrigem In gcolProdutoSRVOrigem
'            If objProdutoSRV.sProdutoSRV = objProdutoSRVOrigem.sProdutoSRV Then
'                If objProdutoSRV.dQuantidade > objProdutoSRVOrigem.dQuantidade Then gError 188038
'                Exit For
'            End If
'        Next
'
'    Next
'
'    Valida_Dados_Tela = SUCESSO
'
'    Exit Function
'
'Erro_Valida_Dados_Tela:
'
'    Valida_Dados_Tela = gErr
'
'    Select Case gErr
'
'        Case 188029
'            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)
'
'        Case 188030
'            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA_GRID1", gErr, iIndice)
'
'        Case 188031
'            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_POSITIVA_GRID", gErr, iIndice)
'
'        Case 188032
'            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_NAO_PREENCHIDA_GRID", gErr, iIndice)
'
'        Case 188033, 188034, 188035, 188036
'
'        Case 188037
'            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_SERVICO_DUPLICADO_GRID", gErr, iIndice, iIndice1)
'
'
'        Case 188038
'            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_MAIOR_ORCADA", gErr, objProdutoSRV.sProdutoSRV, objProdutoSRVOrigem.dQuantidade, objProdutoSRV.dQuantidade)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188039)
'
'    End Select
'
'End Function
'
'Private Function Move_ProdutoSRV_Memoria() As Long
''Move os dados da tela para objGarantia
'
'Dim lErro As Long
'Dim objProdutoSRV As ClassProdutoSRV
'Dim sServicoFormatado As String
'Dim iServicoPreenchido As Integer
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_Move_ProdutoSRV_Memoria
'
'    For iIndice = 1 To gcolProdutoSRVDestino.Count
'        gcolProdutoSRVDestino.Remove (1)
'    Next
'
'    For iIndice = 1 To objGridServico.iLinhasExistentes
'
'        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Servico_Col), sServicoFormatado, iServicoPreenchido)
'        If lErro <> SUCESSO Then gError 188040
'
'        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 188041
'
'
'        Set objProdutoSRV = New ClassProdutoSRV
'
'        objProdutoSRV.sServicoSRV = sServicoFormatado
'        objProdutoSRV.dQuantidade = StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))
'        objProdutoSRV.sProdutoSRV = sProdutoFormatado
'
'        gcolProdutoSRVDestino.Add objProdutoSRV
'
'    Next
'
'    Move_ProdutoSRV_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_ProdutoSRV_Memoria:
'
'    Move_ProdutoSRV_Memoria = gErr
'
'    Select Case gErr
'
'        Case 188040, 188041
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188042)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'
