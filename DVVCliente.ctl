VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl DVVClienteOcx 
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   6795
   Begin VB.CommandButton DVVsCadastradas 
      Caption         =   "DVVs Cadastradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   8
      Top             =   5295
      Width           =   2085
   End
   Begin VB.CommandButton BotaoProdutos 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   9
      Top             =   5310
      Width           =   1365
   End
   Begin VB.CheckBox Paletizacao 
      Height          =   210
      Left            =   4755
      TabIndex        =   7
      Top             =   2895
      Width           =   1305
   End
   Begin VB.TextBox DescricaoProduto 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   1575
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2910
      Width           =   2145
   End
   Begin VB.ComboBox Frete 
      Height          =   315
      Left            =   1530
      TabIndex        =   2
      Top             =   1560
      Width           =   2190
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4380
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "DVVCliente.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "DVVCliente.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "DVVCliente.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "DVVCliente.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   975
      Width           =   2190
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   315
      Left            =   1530
      TabIndex        =   0
      Top             =   375
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   2910
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PercDVV 
      Height          =   225
      Left            =   3810
      TabIndex        =   6
      Top             =   2910
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0%"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridDVVCliente 
      Height          =   3015
      Left            =   210
      TabIndex        =   3
      Top             =   2160
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5318
      _Version        =   393216
   End
   Begin VB.Label LabelFrete 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Frete:"
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
      TabIndex        =   17
      Top             =   1605
      Width           =   1215
   End
   Begin VB.Label LabelFilial 
      AutoSize        =   -1  'True
      Caption         =   "Filial:"
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
      Left            =   990
      TabIndex        =   15
      Top             =   1035
      Width           =   480
   End
   Begin VB.Label LabelCliente 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
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
      Left            =   795
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   14
      Top             =   435
      Width           =   660
   End
End
Attribute VB_Name = "DVVClienteOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'evento dos browsers
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoDVVCliente As AdmEvento
Attribute objEventoDVVCliente.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

'controle de alteracao
Dim iAlterado As Integer
Dim iClienteAlterado As Integer

Dim objGrid As AdmGrid
                            
'variaveis do controle do grid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_PercDVV_Col As Integer
Dim iGrid_Paletizacao_Col As Integer

Public Function Trata_Parametros(Optional objDVVCliente As ClassDVVCliente) As Long
'espera o cód do cliente e a filial

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objDVVCliente Is Nothing) Then
        
        'faz uma busca no bd a partir do cód e da filial
        lErro = CF("DVVCliente_Le", objDVVCliente)
        If lErro <> SUCESSO And lErro <> 116481 Then gError 116428
        
        'se achou no BD
        If lErro = SUCESSO Then

            'Coloca o resultado da busca na tela
            lErro = Traz_DVVCliente_Tela(objDVVCliente)
            If lErro <> SUCESSO Then gError 116429

        Else
        
            'Limpa a tela
            Call Limpa_Tela_DVVCLiente
            
            If objDVVCliente.lCodCliente <> 0 Then
                'coloca o codigo do cliente na tela
                Cliente.Text = CStr(objDVVCliente.lCodCliente)
                Call Cliente_Validate(bSGECancelDummy)
            Else
                Cliente.Text = ""
            End If
            
        End If
                
    End If

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 116429, 116428

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159131)

    End Select

    Exit Function

End Function

Public Sub Form_Load()
'Carrega as configurações iniciais da tela

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega a ComboBox Frete com os tipos de frete
    lErro = Carrega_Frete()
    If lErro <> SUCESSO Then gError 116430
    
    'inicializa a mask. do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 116431
    
    'Inicializa o objGrid
    Set objGrid = New AdmGrid
    
    'inicializa os eventos
    Set objEventoProduto = New AdmEvento
    Set objEventoDVVCliente = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    'inicializa o grid
    lErro = Inicializa_Grid_DVVCliente(objGrid)
    If lErro <> SUCESSO Then gError 116432

    'zera as variáveis de alteração
    iAlterado = 0
    iClienteAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 116430, 116431, 116432

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159132)

    End Select
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_DVVCliente(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Paletização")

   'campos de edição do grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (PercDVV.Name)
    objGridInt.colCampo.Add (Paletizacao.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Produto_Col = 1
    iGrid_DescricaoProduto_Col = 2
    iGrid_PercDVV_Col = 3
    iGrid_Paletizacao_Col = 4

    'passa o grid p/ o obj
    objGridInt.objGrid = GridDVVCliente
    
    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PRODUTOS_DVVCLIENTE + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10

    'largura da 1ª coluna
    GridDVVCliente.ColWidth(0) = 400

    'largura Manual das demias colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_DVVCliente = SUCESSO

End Function

Private Function Carrega_Frete() As Long
'Carrega a combobox Filial
    
Dim lErro As Long
Dim objCodigoFrete As New AdmCodigoNome
Dim colCodigoFrete As New AdmColCodigoNome
    
On Error GoTo Erro_Carrega_Frete
    
    'Leitura dos códigos e nome dos Fretes
    lErro = CF("Cod_Nomes_Le", "TipoFreteFP", "Codigo", "NomeReduzido", STRING_TIPO_FRETE_NOME_REDUZIDO, colCodigoFrete)
    If lErro <> SUCESSO Then gError 116433
    
    'Preenche a ComboBox Frete com código e nome dos fretes
    For Each objCodigoFrete In colCodigoFrete
        Frete.AddItem objCodigoFrete.iCodigo & SEPARADOR & objCodigoFrete.sNome
        Frete.ItemData(Frete.NewIndex) = objCodigoFrete.iCodigo
    Next
    
    Carrega_Frete = SUCESSO
    
    Exit Function

Erro_Carrega_Frete:

    Carrega_Frete = gErr

    Select Case gErr

        Case 116433

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159133)

    End Select

    Exit Function

End Function

Private Sub Cliente_Change()
        iClienteAlterado = REGISTRO_ALTERADO
        iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DVVsCadastradas_Click()

Dim objDVVCliente As New ClassDVVCliente
Dim colSelecao As New Collection

    'preenche c/ o codigo do cliente da tela
    objDVVCliente.lCodCliente = Codigo_Extrai(Cliente.Text)

    'preenche c/ o codigo da filial na tela
    objDVVCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    
    Call Chama_Tela("DVVClienteLista", colSelecao, objDVVCliente, objEventoDVVCliente)

End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Frete_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o codigo do Cliente
    objcliente.lCodigo = Codigo_Extrai(Cliente.Text)

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub BotaoGravar_Click()
'inicia a gravacao

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 116434

    'Limpa a Tela
    Call Limpa_Tela_DVVCLiente

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116434

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159134)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objDVVCliente As New ClassDVVCliente

On Error GoTo Erro_Gravar_Registro

    'transforma o ponteiro um ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'verifica se o cliente está preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 116435
    
    'verifica se a filial está preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 116436
   
    'Se não existir linhas no grid ==> erro
'    If objGrid.iLinhasExistentes = 0 Then gError 116437
   
    'carrega o obj
    lErro = Move_Tela_Memoria(objDVVCliente)
    If lErro <> SUCESSO Then gError 116438

    'Grava no BD
    lErro = CF("DVVCliente_Grava", objDVVCliente)
    If lErro <> SUCESSO Then gError 116439

    'volta o ponteiro no padrao
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 116438, 116439
        
        Case 116436
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 116435
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 116437
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159135)

    End Select

    Exit Function

End Function

Private Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
    
    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridDVVCliente.Row = 0 Then gError 116440

        sProduto1 = GridDVVCliente.TextMatrix(GridDVVCliente.Row, iGrid_Produto_Col)

    End If

    'formata o produto
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 116441

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 116440
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 116441

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159136)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoDVVCliente_evSelecao(obj1 As Object)
'preenche a tela c/ os dados selecionados no browser

Dim objDVVCliente As ClassDVVCliente
Dim lErro As Long

On Error GoTo Erro_objEventoDVVCliente_evSelecao

    Set objDVVCliente = obj1

    'traz os dados p/ a tela
    lErro = Traz_DVVCliente_Tela(objDVVCliente)
    If lErro <> SUCESSO Then gError 116442

    Me.Show

    Exit Sub

Erro_objEventoDVVCliente_evSelecao:

    Select Case gErr
    
        Case 116442
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159137)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim sDescricao As String
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridDVVCliente.Row < 1 Then Exit Sub

    'formata o produto
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 116443

    'inclui no controle
    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True

    'inclui no controle do grid
    GridDVVCliente.TextMatrix(GridDVVCliente.Row, iGrid_Produto_Col) = Produto.Text

    'Faz o Tratamento do produto
    lErro = Produto_Saida_Celula(sDescricao)
    If lErro <> SUCESSO Then gError 116444
    
    'coloca a descricao do prod. no grid
    GridDVVCliente.TextMatrix(GridDVVCliente.Row, iGrid_DescricaoProduto_Col) = sDescricao
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
            
        Case 116443
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 116444

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159138)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim objcliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate
   
    If iClienteAlterado <> 0 Then

        Filial.Clear
        
        'Se o Cliente foi preenchido
        If Len(Trim(Cliente.Text)) > 0 Then
    
            'Busca o Cliente no BD
            lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
            If lErro <> SUCESSO And lErro <> 6668 Then gError 116445
        
            Cliente.Text = objcliente.sNomeReduzido
        
            lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 116446

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)
    
        End If
        
    End If

    iClienteAlterado = 0
    
    Exit Sub
        
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
    
        Case 116445, 116446
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159139)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer

On Error GoTo Erro_Filial_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'verifica se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 119522

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 116447

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objFilialCliente.iCodFilial = iCodigo

        'Tentativa de leitura da Filial com esse código no BD
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 116448

        If lErro = 12567 Then gError 119521  'Não encontrou Filial no  BD

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = objFilialCliente.sNome

    End If
        
    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 116449

    Exit Sub

Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 116447, 116448

        Case 119522
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 116449, 119521
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159140)

    End Select

    Exit Sub

End Sub

Private Sub Frete_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objDVVCliente As New ClassDVVCliente
Dim iCodigo As Integer

On Error GoTo Erro_Frete_Validate

    'Verifica se foi preenchida a ComboBox Frete
    If Len(Trim(Frete.Text)) = 0 Then Exit Sub
    
    'Verifica se está preenchida com o ítem selecionado na ComboBox Frete
    If Frete.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Frete, iCodigo)
    If lErro <> SUCESSO And lErro <> 6731 Then gError 116450
        
    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 116451

    Exit Sub

Erro_Frete_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 116450

        Case 116451
            Call Rotina_Erro(vbOKOnly, "ERRO_FRETE_NAO_ENCONTRADO", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159141)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.lCodigo
    
    'Dispara o Validate de Cliente
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 116452

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 116452

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159142)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long, objDVVCliente As New ClassDVVCliente

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "DVVCliente"

    lErro = Move_Tela_Memoria(objDVVCliente)
    If lErro <> SUCESSO Then gError 106644
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodCliente", objDVVCliente.lCodCliente, 0, "CodCliente"
    colCampoValor.Add "CodFilial", objDVVCliente.iCodFilial, 0, "CodFilial"
    colCampoValor.Add "TipoFrete", objDVVCliente.iTipoFrete, 0, "TipoFrete"

    'adiciona FilialEmpresa
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 106644
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159143)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objDVVCliente As ClassDVVCliente) As Long
'Move os dados da tela para objCategoriaProduto

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim objDVVClienteProd As ClassDVVClienteProd
Dim objcliente As New ClassCliente

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Cliente.Text)) > 0 Then
    
        'Lê o Cliente a partir do Nome Reduzido
        objcliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 106645
        
        'Se não econtrou o Cliente, erro
        If lErro = 12348 Then gError 106646
        
        objDVVCliente.lCodCliente = objcliente.lCodigo
    
    End If
    
    If Len(Trim(Filial.Text)) > 0 Then objDVVCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    If Len(Trim(Frete.Text)) > 0 Then objDVVCliente.iTipoFrete = Codigo_Extrai(Frete.Text)
    objDVVCliente.iFilialEmpresa = giFilialEmpresa

    'preenche uma colecao com todas as linhas "existentes" do grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        Set objDVVClienteProd = New ClassDVVClienteProd

        'formata o produto
        lErro = CF("Produto_Formata", GridDVVCliente.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 116453
        
        'Se o produto não estiver preenchido => erro
        If iPreenchido = PRODUTO_VAZIO Then gError 116454
        
        'preenche o obj c/ o produto formatado
        objDVVClienteProd.sProduto = sProduto
        
        'verifica se o % está preenchido
        If Len(Trim(GridDVVCliente.TextMatrix(iIndice, iGrid_PercDVV_Col))) = 0 Then gError 116493
        
        'carrega o pecentual no obj
        objDVVClienteProd.dPercDVV = PercentParaDbl(GridDVVCliente.TextMatrix(iIndice, iGrid_PercDVV_Col))
        
        'verifica se o produto é paletizado
        If StrParaInt(GridDVVCliente.TextMatrix(iIndice, iGrid_Paletizacao_Col)) = MARCADO Then
            objDVVClienteProd.iPaletizacao = MARCADO
        Else
            objDVVClienteProd.iPaletizacao = DESMARCADO
        End If
        
        'Verifica se já existe o produto na coleção
        For iIndice1 = 1 To objDVVCliente.colDVVCLienteProd.Count
            If UCase(objDVVClienteProd.sProduto) = UCase(objDVVCliente.colDVVCLienteProd.Item(iIndice1).sProduto) Then gError 116455
        Next
    
        'adiciona o obj na col.
        objDVVCliente.colDVVCLienteProd.Add objDVVClienteProd

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr

        Case 106646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
        
        Case 116453, 106645

        Case 116454
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PRODUTO_GRID", gErr, iIndice)

        Case 116455
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_NO_GRID", gErr, objDVVClienteProd.sProduto, iIndice)
        
        Case 116493
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENT_NAO_INFORMADO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159144)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'inicia a etapa de exclusão de registros

Dim lErro As Long
Dim objDVVCliente As New ClassDVVCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'ponteiro p/ ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o cliente foi informado
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 116457

    'verifica se a filial foi informada
    If Len(Trim(Filial.Text)) = 0 Then gError 116458
    
    'pergunta se deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_DVVCLIENTE", Cliente.Text, Filial.Text)
    
    'se não, sai da rotia
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'carrega o obj p/ ser passado como parametro
    lErro = Move_Tela_Memoria(objDVVCliente)
    If lErro <> SUCESSO Then gError 106647
        
    'Faz a exclusão do DVVCliente
    lErro = CF("DVVCliente_Exclui", objDVVCliente)
    If lErro <> SUCESSO Then gError 116459

    'Limpa a Tela
    Call Limpa_Tela_DVVCLiente
    
    'ampulheta p/ padrão
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 116457
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 116458
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 116459, 106647
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159145)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objDVVCliente As New ClassDVVCliente

On Error GoTo Erro_Tela_Preenche

    'preenche o obj p/ ser passado como parametro
    objDVVCliente.iCodFilial = colCampoValor.Item("CodFilial").vValor
    objDVVCliente.lCodCliente = colCampoValor.Item("CodCliente").vValor
    objDVVCliente.iTipoFrete = colCampoValor.Item("TipoFrete").vValor
    objDVVCliente.iFilialEmpresa = giFilialEmpresa

    'Traz dados da Categoria para a Tela
    lErro = Traz_DVVCliente_Tela(objDVVCliente)
    If lErro <> SUCESSO Then gError 116460

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 116460

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159146)

    End Select

    Exit Sub

End Sub

Private Function Traz_DVVCliente_Tela(objDVVCliente As ClassDVVCliente) As Long
'traz os dados do Bd p/ a tela

Dim lErro As Long
Dim iIndice As Integer
Dim objDVVClienteProd As ClassDVVClienteProd
Dim sProdutoEnxuto As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Traz_DVVCliente_Tela
        
    'Exibe o cliente na tela
    Cliente.Text = objDVVCliente.lCodCliente
    Call Cliente_Validate(bSGECancelDummy)

    'Exibe a filial na tela
    Filial.Text = objDVVCliente.iCodFilial
    Call Filial_Validate(bSGECancelDummy)
    
    'verifica se o frete foi preenchido
    If objDVVCliente.iTipoFrete <> 0 Then
        'Exibe o frete na tela
        Frete.Text = objDVVCliente.iTipoFrete
        Call Frete_Validate(bSGECancelDummy)
    Else
        Frete.Text = ""
    End If

    'Lê a tabela DVVClienteProd p/ buscar os produtos a partir do cliente, filial
    lErro = CF("DVVCliente_Le_Itens", objDVVCliente)
    If lErro <> SUCESSO And lErro <> 116486 Then gError 116461

    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGrid)

    'Exibe os dados da coleção na tela
    For Each objDVVClienteProd In objDVVCliente.colDVVCLienteProd
        
        iIndice = iIndice + 1
        
        'formata o produto
        lErro = Mascara_RetornaProdutoEnxuto(objDVVCliente.colDVVCLienteProd.Item(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 116462

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
                
        'Insere o produto no Grid DVVCliente
        GridDVVCliente.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        
'        'busca a descricao do produto
'        lErro = Produto_Saida_Celula(sDescricao)
'        If lErro <> SUCESSO Then gError 116463
        
        objProduto.sCodigo = objDVVCliente.colDVVCLienteProd.Item(iIndice).sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 166494
        
        'coloca a descrição do produto no grid
        GridDVVCliente.TextMatrix(iIndice, iGrid_DescricaoProduto_Col) = objProduto.sDescricao
        
        'coloca o percentual no grid
        GridDVVCliente.TextMatrix(iIndice, iGrid_PercDVV_Col) = Format((objDVVCliente.colDVVCLienteProd.Item(iIndice).dPercDVV), "Percent")
        
        'verifica se é paletizado
        If objDVVClienteProd.iPaletizacao = MARCADO Then
            GridDVVCliente.TextMatrix(iIndice, iGrid_Paletizacao_Col) = MARCADO
        Else
            GridDVVCliente.TextMatrix(iIndice, iGrid_Paletizacao_Col) = DESMARCADO
        End If
    
    Next

    objGrid.iLinhasExistentes = iIndice

    'da um "refresh" nas checkbox do grid
    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 166494

    'zera as variaveis de alteração
    iAlterado = 0
    iClienteAlterado = 0

    Traz_DVVCliente_Tela = SUCESSO

    Exit Function

Erro_Traz_DVVCliente_Tela:

    Traz_DVVCliente_Tela = gErr

    Select Case gErr

        Case 116463, 116462, 116461, 166494

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159147)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'botão que limpa a tela

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 116464

    'Limpa a Tela
    Call Limpa_Tela_DVVCLiente

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 116464

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159148)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_DVVCLiente()
'sub que limpa a tela
    
On Error GoTo Erro_Limpa_Tela_DVVCLiente
    
    'limpa as textbox
    Call Limpa_Tela(Me)
    
    'limpa o grid
    Call Grid_Limpa(objGrid)
    
    'limpa as combos
    Filial.Clear
    Frete.Text = ""

    'zera as variaveis de alteracao
    iClienteAlterado = 0
    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_DVVCLiente:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159149)

    End Select
    
    Exit Sub

End Sub

Sub GridDVVCliente_Click()
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If
    
End Sub

Sub GridDVVCliente_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Sub GridDVVCliente_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Sub GridDVVCliente_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Sub GridDVVCliente_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridDVVCliente_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    Exit Sub

Erro_GridDVVCliente_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159150)

    End Select

    Exit Sub

End Sub

Sub GridDVVCliente_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Sub GridDVVCliente_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Sub GridDVVCliente_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Sub GridDVVCliente_Scroll()
    Call Grid_Scroll(objGrid)
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PercDVV_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercDVV_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercDVV_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub

Private Sub PercDVV_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub

Private Sub PercDVV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = PercDVV
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Paletizacao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Paletizacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub

Private Sub Paletizacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub

Private Sub Paletizacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Paletizacao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz o tratamento de saida de célula

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    'Sucesso => ...
    If lErro = SUCESSO Then
        
        Select Case GridDVVCliente.Col

            Case iGrid_Produto_Col
                'faz a saida da celula do produto
                lErro = Saida_Celula_Produto(objGridInt)
                If lErro <> SUCESSO Then gError 116465

            Case iGrid_PercDVV_Col
                'faz a saida da celula do percentual
                lErro = Saida_Celula_PercDVV(objGridInt)
                If lErro <> SUCESSO Then gError 116466
            
            Case iGrid_Paletizacao_Col
                'faz a saida da celula da paletizacao
                lErro = Saida_Celula_Paletizacao(objGridInt)
                If lErro <> SUCESSO Then gError 116467

        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 116468
    
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr

        Case 116465 To 116467
        
        Case 116468
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159151)
    
    End Select
    
    Exit Function

End Function

Public Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'faz o tratamento de saida de célula do produto

Dim lErro As Long
Dim sDescricao As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
    
    'verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then
        
        'busca a validação do produto e a descricao
        lErro = Produto_Saida_Celula(sDescricao)
        If lErro <> SUCESSO Then gError 116469
        
    End If
        
    'coloca a descrição do produto
    GridDVVCliente.TextMatrix(GridDVVCliente.Row, iGrid_DescricaoProduto_Col) = sDescricao
    
    'p/ o rotina_grid_enable
    GridDVVCliente.TextMatrix(GridDVVCliente.Row, iGrid_Produto_Col) = ""
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116470

    Saida_Celula_Produto = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr
    
    Select Case gErr
    
        Case 116469
    
        Case 116470
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159152)
    
    End Select
    
    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)
'habilita / desabilita o campo produto

Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim lErro As Long
        
On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid
    lErro = CF("Produto_Formata", GridDVVCliente.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 116492

    Select Case objControl.Name

        Case Produto.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If
    
    End Select
    
    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 116492
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159153)
    
    End Select
    
    Exit Sub
    
End Sub

Function Produto_Saida_Celula(sDescricao As String) As Long
'faz a validação do produto no grid

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Produto_Saida_Celula

    'limpa a descrição
    sDescricao = ""

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 116471
    
    If lErro = 51381 Then gError 116472

    'se o produto existe
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        'retorna o produto enxuto
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 116473

        'coloca o cód. do produto no controle
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
    
        'preenche a Descricao Produto
        sDescricao = objProduto.sDescricao
        
        'se necessário, cria + uma linha
        If GridDVVCliente.Row - GridDVVCliente.FixedRows = objGrid.iLinhasExistentes Then objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1

    End If

    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes
        If iIndice <> GridDVVCliente.Row Then
            If GridDVVCliente.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 116474
        End If
    Next

    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case 116471, 116473

        Case 116472
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGrid)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGrid)
            End If

        Case 116474
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_NO_GRID", gErr, Produto.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159154)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Paletizacao(objGridInt As AdmGrid) As Long
'faz a saida da celula da checkbox paletizacao

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Paletizacao

    Set objGridInt.objControle = Paletizacao
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116475

    Saida_Celula_Paletizacao = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Paletizacao:

    Saida_Celula_Paletizacao = gErr
    
    Select Case gErr
    
        Case 116475
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159155)
    
    End Select
    
    Exit Function

End Function

Public Function Saida_Celula_PercDVV(objGridInt As AdmGrid) As Long
'faz a saida da celula percentual

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PercDVV

    Set objGridInt.objControle = PercDVV
    
    'Se estiver preenchida
    If Len(Trim(PercDVV.Text)) > 0 Then
        'Critica o valor
        lErro = Porcentagem_Critica(PercDVV.Text)
        If lErro <> SUCESSO Then gError 116476
        
        'se necessário, cria + uma linha
        If GridDVVCliente.Row - GridDVVCliente.FixedRows = objGrid.iLinhasExistentes Then objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1

    End If
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116477

    Saida_Celula_PercDVV = SUCESSO
    
    Exit Function

Erro_Saida_Celula_PercDVV:

    Saida_Celula_PercDVV = gErr
    
    Select Case gErr
    
        Case 116476
    
        Case 116477
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159156)
    
    End Select
    
    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        End If
        
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    'libera os objs
    Set objEventoProduto = Nothing
    Set objEventoDVVCliente = Nothing
    Set objEventoCliente = Nothing
    Set objGrid = Nothing
    
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Despesas Variáveis de Venda por Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "DVVCliente"
    
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

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub LabelFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFrete, Source, X, Y)
End Sub

Private Sub LabelFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFrete, Button, Shift, X, Y)
End Sub














'/////////// tarefa de altreracao de categoria produto

''Alterado por Ivan 25/4/03
''inclusão de leitura de novos valores
'Public Function CategoriaProduto_Le_Item(ByVal objCategoriaProdutoItem As ClassCategoriaProdutoItem) As Long
''Le na tabela de CategoriaProdutoItem a Categoria e o Item de uma deteminada Categoria de Produto
''Alterada por tulio em 22/04
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim lComando As Long
'Dim sCategoria As String
'Dim sItem As String
'Dim sDescricaoItem As String
'Dim adValor(0 To 2) As Double
'
'On Error GoTo Erro_CategoriaProduto_Le_Item
'
'    'Inicializar comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 22600
'
'    sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
'    sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
'    sDescricaoItem = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)
'
'    'Executar comando SQL
'    lErro = Comando_Executar(lComando, "SELECT Categoria, Item, Descricao, Valor1, Valor2, Valor3 FROM CategoriaProdutoItem WHERE Categoria = ? AND Item = ?", sCategoria, sItem, sDescricaoItem, adValor(0), adValor(1), adValor(2), objCategoriaProdutoItem.sCategoria, objCategoriaProdutoItem.sItem)
'    If lErro <> AD_SQL_SUCESSO Then Error 22601
'
'    lErro = Comando_BuscarProximo(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22602
'
'    'Se não encontrou
'    If lErro = AD_SQL_SEM_DADOS Then Error 22603
'
'    'Atribui ao obj os dados lidos do bd
'    objCategoriaProdutoItem.sCategoria = sCategoria
'    objCategoriaProdutoItem.sItem = sItem
'    objCategoriaProdutoItem.sDescricao = sDescricaoItem
'    objCategoriaProdutoItem.dvalor1 = adValor(0)
'    objCategoriaProdutoItem.dvalor2 = adValor(1)
'    objCategoriaProdutoItem.dvalor3 = adValor(2)
'
'    Call Comando_Fechar(lComando)
'
'    CategoriaProduto_Le_Item = SUCESSO
'
'    Exit Function
'
'Erro_CategoriaProduto_Le_Item:
'
'    CategoriaProduto_Le_Item = Err
'
'    Select Case Err
'
'        Case 22600
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 22601, 22602
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", Err, objCategoriaProdutoItem.sCategoria, objCategoriaProdutoItem.sItem)
'
'        Case 22603
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159157)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
