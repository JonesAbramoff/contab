VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CategoriaClienteOcx 
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
         Picture         =   "CategoriaClienteOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CategoriaClienteOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CategoriaClienteOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CategoriaClienteOcx.ctx":0816
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
Attribute VB_Name = "CategoriaClienteOcx"
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
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se Categoria foi preenchida
    If Len(Trim(Categoria.Text)) = 0 Then Error 28890
    
    'Preenche objCategoriaCliente
    objCategoriaCliente.sCategoria = Categoria.Text
    
    'Envia aviso perguntando se realmente deseja excluir Categoria
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CATEGORIACLIENTE", objCategoriaCliente.sCategoria)

    If vbMsgRes = vbYes Then

        GL_objMDIForm.MousePointer = vbHourglass
    
        'Exclui Categoria
        lErro = CF("CategoriaCliente_Exclui",objCategoriaCliente)
        If lErro <> SUCESSO Then Error 28891

        GL_objMDIForm.MousePointer = vbDefault
    
        'Exclui a Categoria da Combo
        For iIndice1 = 0 To Categoria.ListCount - 1

            If Categoria.List(iIndice1) = objCategoriaCliente.sCategoria Then

                Categoria.RemoveItem (iIndice1)

                Exit For

            End If

        Next

        'Limpa a Tela
        Call Limpa_Tela_CategoriaCliente

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 28890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_INFORMADA", Err)

        Case 28891

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144253)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 28835

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case Err

        Case 28835

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144254)

    End Select

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 28836

    'Limpa a tela
    Call Limpa_Tela_CategoriaCliente

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 28836

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144255)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 28878

    'Limpa a Tela
    Call Limpa_Tela_CategoriaCliente

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 28878

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144256)

    End Select

End Sub

Sub Limpa_Tela_CategoriaCliente()

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
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Categoria_Click

    iAlterado = REGISTRO_ALTERADO

    'Se alguém estiver selecionado
    If Categoria.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = Categoria.Text
        
        'Se no Trata_Parametros preencheu com um Item passado não renova o Grid
        If iRenova_Grid = 0 Then
                        
            lErro = Traz_CategoriaCliente_Tela(objCategoriaCliente)
            If lErro <> SUCESSO Then Error 28879
        
        End If

        iRenova_Grid = 0
        
    End If

    Exit Sub

Erro_Categoria_Click:

    Select Case Err

        Case 28879 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144257)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Categoria_Validate(Cancel As Boolean)

Dim iIndice As Integer, lErro As Long

On Error GoTo Erro_Categoria_Validate

    If Len(Trim(Categoria.Text)) <> 0 Then
    
        If Categoria.ListIndex = -1 Then
            
            If Len(Trim(Categoria.Text)) > STRING_CATEGORIACLIENTE_CATEGORIA Then Error 19369
            
            'Seleciona na Combo um item igual ao digitado
            Call Combo_Item_Igual(Categoria)
                    
        End If
        
    End If

    Exit Sub
    
Erro_Categoria_Validate:

    Select Case Err

        Case 19369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_TAMMAX", Err, STRING_CATEGORIACLIENTE_CATEGORIA)
            Cancel = True
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144258)

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
    If lErro <> SUCESSO Then Error 28837

    'Inicializa o Grid
    Set objGrid = New AdmGrid
    lErro = Inicializa_Grid_ItensCategoria(objGrid)
    If lErro <> SUCESSO Then Error 28838

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 28837, 28838 ' Tratados nas rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144259)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objCategoriaCliente As ClassCategoriaCliente, Optional sItem As String) As Long

Dim lErro As Long
On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma categoria foi passada por parâmetro
    If Not (objCategoriaCliente Is Nothing) Then
        
        'Traz a Categoria do Cliente para a tela
        lErro = Traz_CategoriaCliente_Tela(objCategoriaCliente)
        If lErro <> SUCESSO And lErro <> 19368 Then Error 28885

        'se a categoria nao está cadastrada
        If lErro = 19368 Then Categoria.Text = objCategoriaCliente.sCategoria
    
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

    Trata_Parametros = Err

    Select Case Err

        Case 28885 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144260)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Function Traz_CategoriaCliente_Tela(objCategoriaCliente As ClassCategoriaCliente) As Long
'Traz os dados da Categoria do Cliente para tela

Dim lErro As Long
Dim iIndice As Integer
Dim colItensCategoria As New Collection

On Error GoTo Erro_Traz_CategoriaCliente_Tela

    'Lê a tabela CategoriaCliente a partir da Categoria
    lErro = CF("CategoriaCliente_Le",objCategoriaCliente)
    If lErro <> SUCESSO And lErro <> 28847 Then Error 28886
    
    If lErro = 28847 Then Error 19368
    
    'Exibe os dados de objCategoriaCliente na tela
    Categoria.Text = objCategoriaCliente.sCategoria
    Descricao.Text = objCategoriaCliente.sDescricao

    'Lê a tabela CategoriaClienteItem à partir da Categoria
    lErro = CF("CategoriaCliente_Le_Itens",objCategoriaCliente, colItensCategoria)
    If lErro <> SUCESSO And lErro <> 28851 Then Error 28888

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

    Traz_CategoriaCliente_Tela = SUCESSO

    Exit Function

Erro_Traz_CategoriaCliente_Tela:

    Traz_CategoriaCliente_Tela = Err

    Select Case Err

        Case 28886, 28888 'Tratados nas Rotinas Chamadas
        
        Case 19368 'Não encontrou --> Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144261)

    End Select

    Exit Function

End Function

Private Function Carrega_Categoria() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaCliente As ClassCategoriaCliente

On Error GoTo Erro_Carrega_Categoria

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriaCliente_Le_Todos",colCategorias)
    If lErro <> SUCESSO Then Error 28839

    For Each objCategoriaCliente In colCategorias

        'Insere na combo Categoria
        Categoria.AddItem objCategoriaCliente.sCategoria

    Next

    Carrega_Categoria = SUCESSO

    Exit Function

Erro_Carrega_Categoria:

    Carrega_Categoria = Err

    Select Case Err

        Case 28839

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144262)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim colItensCategoria As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a Categoria está preenchida
    If Len(Trim(Categoria.Text)) = 0 Then Error 28853

    'Chama Move_Tela_Memoria para passar os dados da tela para  os objetos
    lErro = Move_Tela_Memoria(objCategoriaCliente, colItensCategoria)
    If lErro <> SUCESSO Then Error 28854

    If colItensCategoria.Count = 0 Then Error 56566
    
    lErro = Trata_Alteracao(objCategoriaCliente, objCategoriaCliente.sCategoria)
    If lErro <> SUCESSO Then Error 32281
        
    'Chama a função de gravacao
    lErro = CF("CategoriaCliente_Grava",objCategoriaCliente, colItensCategoria)
    If lErro <> SUCESSO Then Error 28855

    'Exclui ( se existir) da lista de Categoria
    Call ListaCategoria_Exclui(objCategoriaCliente.sCategoria)

    'Adiciona na lista de Categoria
    Categoria.AddItem objCategoriaCliente.sCategoria

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 56566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_SEM_ITEM_CORRESPONDENTE", Err)
        
        Case 28853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_INFORMADA", Err)

        Case 28854, 28855, 32281 'Tratados nas Rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144263)

    End Select

End Function

Function Move_Tela_Memoria(objCategoriaCliente As ClassCategoriaCliente, colItensCategoria As Collection) As Long
'Move os dados da tela para memória

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objCategoriaCliente
    If Len(Trim(Categoria.Text)) > 0 Then objCategoriaCliente.sCategoria = Trim(Categoria.Text)
    If Len(Trim(Descricao.Text)) > 0 Then objCategoriaCliente.sDescricao = Descricao.Text

    'Ir preenchendo uma coleção com todas as linhas "existentes" do grid
        'o item e sua descrição tem que estar preenchidos, senão erro
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a DescricaoItem está preenchida  e o Item não está preenchido
        If Len(Trim(GridItensCategoria.TextMatrix(iIndice, iGrid_DescricaoItem_Col))) > 0 And Len(Trim(GridItensCategoria.TextMatrix(iIndice, iGrid_Item_Col))) = 0 Then Error 28856

        'Verifica se o Item foi preenchido
        If Len(Trim(GridItensCategoria.TextMatrix(iIndice, iGrid_Item_Col))) <> 0 Then

            Set objCategoriaClienteItem = New ClassCategoriaClienteItem

            objCategoriaClienteItem.sCategoria = objCategoriaCliente.sCategoria
            objCategoriaClienteItem.sItem = Trim(GridItensCategoria.TextMatrix(iIndice, iGrid_Item_Col))
            objCategoriaClienteItem.sDescricao = GridItensCategoria.TextMatrix(iIndice, iGrid_DescricaoItem_Col)

            'Verifica se já existe o Item na coleção
            For iIndice1 = 1 To colItensCategoria.Count

                If objCategoriaClienteItem.sItem = colItensCategoria.Item(iIndice1).sItem Then Error 28857

            Next
            
            'Adiciona na colecao
            colItensCategoria.Add objCategoriaClienteItem

        End If

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 28856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_ITEM_CATEGORIACLIENTE", Err)

        Case 28857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_REPETIDO_NO_GRID", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144264)

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
                If lErro <> SUCESSO Then Error 28880

            Case iGrid_DescricaoItem_Col
                
                'Critica a Saída da Descrição do Item
                lErro = Saida_Celula_DescricaoItem(objGridInt)
                If lErro <> SUCESSO Then Error 28881

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 28882

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 28880, 28881 'Tratados nas Rotinas Chamadas

        Case 28882
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144265)

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
    If lErro <> SUCESSO Then Error 28883

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = Err

    Select Case Err

        Case 28883
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144266)

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
    If lErro <> SUCESSO Then Error 28884

    Saida_Celula_Item = SUCESSO

    Exit Function

Erro_Saida_Celula_Item:

    Saida_Celula_Item = Err

    Select Case Err

        Case 28884
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144267)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CategoriaCliente"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Categoria", Categoria.Text, STRING_CATEGORIACLIENTE_CATEGORIA, "Categoria"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144268)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Tela_Preenche

    objCategoriaCliente.sCategoria = colCampoValor.Item("Categoria").vValor

    If Len(Trim(objCategoriaCliente.sCategoria)) > 0 Then

       'Traz dados da Categoria para a Tela
        lErro = Traz_CategoriaCliente_Tela(objCategoriaCliente)
        If lErro <> SUCESSO Then Error 28915

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 28915 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144269)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CATEGORIAS_CLIENTE
    Set Form_Load_Ocx = Me
    Caption = "Categorias de Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CategoriaCliente"
    
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

