VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CategoriaFornecOcx 
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   LockControls    =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   6660
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4410
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CategoriaFornecOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CategoriaFornecOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CategoriaFornecOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CategoriaFornecOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Categoria 
      Height          =   315
      Left            =   1110
      TabIndex        =   0
      Top             =   360
      Width           =   2610
   End
   Begin MSMask.MaskEdBox DescricaoItem 
      Height          =   225
      Left            =   2235
      TabIndex        =   3
      Top             =   1860
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Item 
      Height          =   225
      Left            =   495
      TabIndex        =   2
      Top             =   1860
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   945
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridItensCategoria 
      Height          =   2070
      Left            =   120
      TabIndex        =   4
      Top             =   1740
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   3651
      _Version        =   393216
      Rows            =   8
      Cols            =   5
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   165
      TabIndex        =   10
      Top             =   420
      Width           =   885
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   11
      Top             =   990
      Width           =   930
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Valores Possíveis"
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
      Left            =   150
      TabIndex        =   12
      Top             =   1500
      Width           =   1530
   End
End
Attribute VB_Name = "CategoriaFornecOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim objGrid As AdmGrid
Dim iGrid_Item_Col As Integer
Dim iGrid_DescricaoItem_Col As Integer
Dim iRenova_Grid As Integer

Private Function Inicializa_Grid_ItensCategoria(objGridInt As AdmGrid) As Long

    'Tela em questão
    Set objGridInt.objForm = Me

    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Descrição Item")

   'Campos de edição do grid
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Item_Col = 1
    iGrid_DescricaoItem_Col = 2

    objGridInt.objGrid = GridItensCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_CATEGORIA + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    GridItensCategoria.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ItensCategoria = SUCESSO

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se Categoria foi preenchida
    If Len(Trim(Categoria.Text)) = 0 Then gError 90526

    'Preenche objCategoriaFornecedor
    objCategoriaFornecedor.sCategoria = Categoria.Text

    'Envia aviso perguntando se realmente deseja excluir Categoria
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CATEGORIAFORNECEDOR", objCategoriaFornecedor.sCategoria)

    If vbMsgRes = vbYes Then

        GL_objMDIForm.MousePointer = vbHourglass

        'Exclui Categoria
        lErro = CF("CategoriaFornecedor_Exclui", objCategoriaFornecedor)
        If lErro <> SUCESSO Then gError 90527

        GL_objMDIForm.MousePointer = vbDefault

        'Exclui a Categoria da Combo
        For iIndice1 = 0 To Categoria.ListCount - 1

            If Categoria.List(iIndice1) = objCategoriaFornecedor.sCategoria Then

                Categoria.RemoveItem (iIndice1)

                Exit For

            End If

        Next

        'Limpa a Tela
        Call Limpa_Tela_CategoriaFornecedor

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 90526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDOR_NAO_INFORMADA", gErr)

        Case 90527

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144270)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 90528

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 90528

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144271)

    End Select

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 90529

    'Limpa a tela
    Call Limpa_Tela_CategoriaFornecedor

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 90529

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144272)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 90530

    'Limpa a Tela
    Call Limpa_Tela_CategoriaFornecedor

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 90530

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144273)

    End Select

End Sub

Sub Limpa_Tela_CategoriaFornecedor()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa os textos das Combos
    Categoria.Text = ""

    'Limpa GridItensCategoria
    Call Grid_Limpa(objGrid)

    iAlterado = 0

End Sub

Private Sub Categoria_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Categoria_Click()

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor

On Error GoTo Erro_Categoria_Click

    iAlterado = REGISTRO_ALTERADO

    'Se alguém estiver selecionado
    If Categoria.ListIndex <> -1 Then

        objCategoriaFornecedor.sCategoria = Categoria.Text

        'Se no Trata_Parametros preencheu com um Item passado não renova o Grid
        If iRenova_Grid = 0 Then

            lErro = Traz_CategoriaFornecedor_Tela(objCategoriaFornecedor)
            If lErro <> SUCESSO Then gError 90531

        End If

        iRenova_Grid = 0

    End If

    Exit Sub

Erro_Categoria_Click:

    Select Case gErr

        Case 90531 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144274)

    End Select

    Exit Sub

End Sub

Private Sub Categoria_Validate(Cancel As Boolean)

Dim iIndice As Integer, lErro As Long

On Error GoTo Erro_Categoria_Validate

    If Len(Trim(Categoria.Text)) <> 0 Then

        If Categoria.ListIndex = -1 Then

            If Len(Trim(Categoria.Text)) > STRING_CATEGORIAFORNECEDOR_CATEGORIA Then gError 90532

            'Seleciona na Combo um item igual ao digitado
            Call Combo_Item_Igual_CI(Categoria)

        End If

    End If

    Exit Sub

Erro_Categoria_Validate:

    Select Case gErr

        Case 90532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDOR_TAMMAX", gErr, STRING_CATEGORIAFORNECEDOR_CATEGORIA)
            Cancel = True

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144275)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub DescricaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()
Dim lErro As Long
On Error GoTo Erro_Form_Load

    'Carrega a ComboBox Categoria com os Códigos
    lErro = Carrega_Categoria()
    If lErro <> SUCESSO Then gError 90533

    'Inicializa o Grid
    Set objGrid = New AdmGrid
    lErro = Inicializa_Grid_ItensCategoria(objGrid)
    If lErro <> SUCESSO Then gError 90534

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 90533, 90534 ' Tratados nas rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144276)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCategoriaFornecedor As ClassCategoriaFornecedor, Optional sItem As String) As Long

Dim lErro As Long
On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma categoria foi passada por parâmetro
    If Not (objCategoriaFornecedor Is Nothing) Then

        'Traz a Categoria do Fornecedor para a tela
        lErro = Traz_CategoriaFornecedor_Tela(objCategoriaFornecedor)
        If lErro <> SUCESSO And lErro <> 90537 Then gError 90535

        'se a categoria nao está cadastrada
        If lErro = 90537 Then Categoria.Text = objCategoriaFornecedor.sCategoria

    End If

    If Len(Trim(sItem)) > 0 Then

        'Caso tenha passado um item como Parametro --> Preenche no Grid
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        GridItensCategoria.TextMatrix(objGrid.iLinhasExistentes, iGrid_Item_Col) = sItem

        'Variavel setada para que no evento click ele não renove o Grid
        iRenova_Grid = 1

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 90535 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144277)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Traz_CategoriaFornecedor_Tela(objCategoriaFornecedor As ClassCategoriaFornecedor) As Long
'Traz os dados da Categoria do Fornecedor para tela

Dim lErro As Long
Dim iIndice As Integer
Dim colItensCategoria As New Collection
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem

On Error GoTo Erro_Traz_CategoriaFornecedor_Tela

    'Lê a tabela CategoriaFornecedor a partir da Categoria
    lErro = CF("CategoriaFornecedor_Le", objCategoriaFornecedor)
    If lErro <> SUCESSO And lErro <> 90592 Then gError 90536

    If lErro = 90592 Then gError 90537

    'Exibe os dados de objCategoriaFornecedor na tela
    Categoria.Text = objCategoriaFornecedor.sCategoria
    Descricao.Text = objCategoriaFornecedor.sDescricao

    'Guarda no obj o nome da categoria que deve ter os itens lidos
    objCategoriaFornecedorItem.sCategoria = objCategoriaFornecedor.sCategoria
    
    'Lê a tabela CategoriaFornecedorItem à partir da Categoria
    lErro = CF("CategoriaFornecedor_Le_Itens", objCategoriaFornecedorItem, colItensCategoria)
    If lErro <> SUCESSO And lErro <> 91180 Then gError 90538

    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGrid)

    'Exibe os dados da coleção na tela
    For iIndice = 1 To colItensCategoria.Count
        'Insere item no Grid
        GridItensCategoria.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = colItensCategoria.Item(iIndice).sDescricao
        GridItensCategoria.TextMatrix(iIndice, iGrid_Item_Col) = colItensCategoria.Item(iIndice).sItem
    Next

    objGrid.iLinhasExistentes = colItensCategoria.Count

    'Zerar iAlterado
    iAlterado = 0

    Traz_CategoriaFornecedor_Tela = SUCESSO

    Exit Function

Erro_Traz_CategoriaFornecedor_Tela:

    Traz_CategoriaFornecedor_Tela = gErr

    Select Case gErr

        Case 90536, 90538 'Tratados nas Rotinas Chamadas

        Case 90537 'Não encontrou --> Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144278)

    End Select

    Exit Function

End Function

Private Function Carrega_Categoria() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaFornecedor As ClassCategoriaFornecedor

On Error GoTo Erro_Carrega_Categoria

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaFornecedor_Le_Todos", colCategorias)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 90539

    For Each objCategoriaFornecedor In colCategorias

        'Insere na combo Categoria
        Categoria.AddItem objCategoriaFornecedor.sCategoria

    Next

    Carrega_Categoria = SUCESSO

    Exit Function

Erro_Carrega_Categoria:

    Carrega_Categoria = gErr

    Select Case gErr

        Case 90539

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144279)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem
Dim colItensCategoria As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se a Categoria está preenchida
    If Len(Trim(Categoria.Text)) = 0 Then gError 90540

    'Chama Move_Tela_Memoria para passar os dados da tela para  os objetos
    lErro = Move_Tela_Memoria(objCategoriaFornecedor, colItensCategoria)
    If lErro <> SUCESSO Then gError 90541

    If colItensCategoria.Count = 0 Then gError 90542

    lErro = Trata_Alteracao(objCategoriaFornecedor, objCategoriaFornecedor.sCategoria)
    If lErro <> SUCESSO Then gError 90543

    'Chama a função de gravacao
    lErro = CF("CategoriaFornecedor_Grava", objCategoriaFornecedor, colItensCategoria)
    If lErro <> SUCESSO Then gError 90544

    'Exclui ( se existir) da lista de Categoria
    Call ListaCategoria_Exclui(objCategoriaFornecedor.sCategoria)

    'Adiciona na lista de Categoria
    Categoria.AddItem objCategoriaFornecedor.sCategoria

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 90542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_SEM_ITEM_CORRESPONDENTE", gErr)

        Case 90540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDOR_NAO_INFORMADA", gErr)

        Case 90541, 90544, 90543 'Tratados nas Rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144280)

    End Select

End Function

Function Move_Tela_Memoria(objCategoriaFornecedor As ClassCategoriaFornecedor, colItensCategoria As Collection) As Long
'Move os dados da tela para memória

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objCategoriaFornecedorItem As ClassCategoriaFornItem

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objCategoriaFornecedor
    If Len(Trim(Categoria.Text)) > 0 Then objCategoriaFornecedor.sCategoria = Trim(Categoria.Text)
    If Len(Trim(Descricao.Text)) > 0 Then objCategoriaFornecedor.sDescricao = Descricao.Text

    'Ir preenchendo uma coleção com todas as linhas "existentes" do grid
        'o item e sua descrição tem que estar preenchidos, senão erro
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a DescricaoItem está preenchida  e o Item não está preenchido
        If Len(Trim(GridItensCategoria.TextMatrix(iIndice, iGrid_DescricaoItem_Col))) > 0 And Len(Trim(GridItensCategoria.TextMatrix(iIndice, iGrid_Item_Col))) = 0 Then gError 90545

        'Verifica se o Item foi preenchido
        If Len(Trim(GridItensCategoria.TextMatrix(iIndice, iGrid_Item_Col))) <> 0 Then

            Set objCategoriaFornecedorItem = New ClassCategoriaFornItem

            objCategoriaFornecedorItem.sCategoria = objCategoriaFornecedor.sCategoria
            objCategoriaFornecedorItem.sItem = Trim(GridItensCategoria.TextMatrix(iIndice, iGrid_Item_Col))
            objCategoriaFornecedorItem.sDescricao = GridItensCategoria.TextMatrix(iIndice, iGrid_DescricaoItem_Col)

            'Verifica se já existe o Item na coleção
            For iIndice1 = 1 To colItensCategoria.Count

                If UCase(objCategoriaFornecedorItem.sItem) = UCase(colItensCategoria.Item(iIndice1).sItem) Then gError 90546

            Next

            'Adiciona na colecao
            colItensCategoria.Add objCategoriaFornecedorItem

        End If

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 90545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_ITEM_CATEGORIAFORNECEDOR", gErr)

        Case 90546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_REPETIDO_NO_GRID", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144281)

    End Select

    Exit Function

End Function

Private Sub ListaCategoria_Exclui(sCategoria As String)
'Exclui a Categoria da lista

Dim iIndice As Integer

    For iIndice = 0 To Categoria.ListCount - 1

        If Categoria.List(iIndice) = sCategoria Then

            Categoria.RemoveItem (iIndice)

            Exit For

        End If

    Next

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, 1, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGrid = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub GridItensCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If

End Sub

Private Sub GridItensCategoria_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridItensCategoria_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridItensCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer

    iLinhasExistentesAnterior = objGrid.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    If iLinhasExistentesAnterior <> objGrid.iLinhasExistentes Then

        iAlterado = REGISTRO_ALTERADO

    End If

End Sub

Private Sub GridItensCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then

        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If

End Sub

Private Sub GridItensCategoria_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridItensCategoria_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridItensCategoria_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridItensCategoria_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub Item_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Item_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Item
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridItensCategoria.Col

            Case iGrid_Item_Col

                'Critica a Saída do Item
                lErro = Saida_Celula_Item(objGridInt)
                If lErro <> SUCESSO Then gError 90547

            Case iGrid_DescricaoItem_Col

                'Critica a Saída da Descrição do Item
                lErro = Saida_Celula_DescricaoItem(objGridInt)
                If lErro <> SUCESSO Then gError 90548

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 90549

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 90547, 90548 'Tratados nas Rotinas Chamadas

        Case 90549
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144282)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DescricaoItem do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    If Len(DescricaoItem.Text) > 0 Then

        If GridItensCategoria.Row - GridItensCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            'Adiciona uma Linha ao Grid
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 90550

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = gErr

    Select Case gErr

        Case 90550
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144283)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Item(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Item

    Set objGridInt.objControle = Item

    If Len(Item.Text) > 0 Then

        If GridItensCategoria.Row - GridItensCategoria.FixedRows = objGridInt.iLinhasExistentes Then

            'Adiciona uma linha ao Grid
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 90551

    Saida_Celula_Item = SUCESSO

    Exit Function

Erro_Saida_Celula_Item:

    Saida_Celula_Item = gErr

    Select Case gErr

        Case 90551
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144284)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CategoriaFornecedor"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Categoria", Categoria.Text, STRING_CATEGORIAFORNECEDOR_CATEGORIA, "Categoria"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144285)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor

On Error GoTo Erro_Tela_Preenche

    objCategoriaFornecedor.sCategoria = colCampoValor.Item("Categoria").vValor

    If Len(Trim(objCategoriaFornecedor.sCategoria)) > 0 Then

       'Traz dados da Categoria para a Tela
        lErro = Traz_CategoriaFornecedor_Tela(objCategoriaFornecedor)
        If lErro <> SUCESSO Then gError 90552

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 90552 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144286)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_CATEGORIAS_FORNECEDOR
    Set Form_Load_Ocx = Me
    Caption = "Categorias de Fornecedores"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CategoriaFornec"

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

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

