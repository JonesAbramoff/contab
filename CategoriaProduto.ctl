VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl CategoriaProduto 
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10365
   ScaleHeight     =   4170
   ScaleWidth      =   10365
   Begin MSMask.MaskEdBox Valor3 
      Height          =   225
      Left            =   8580
      TabIndex        =   7
      Top             =   2205
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Valor2 
      Height          =   225
      Left            =   7350
      TabIndex        =   6
      Top             =   2205
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Valor1 
      Height          =   225
      Left            =   6105
      TabIndex        =   5
      Top             =   2250
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptChar      =   " "
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8085
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   180
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CategoriaProduto.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CategoriaProduto.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CategoriaProduto.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CategoriaProduto.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Categoria 
      Height          =   315
      Left            =   1260
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   330
      Width           =   2610
   End
   Begin MSMask.MaskEdBox DescricaoItem 
      Height          =   225
      Left            =   2475
      TabIndex        =   4
      Top             =   2235
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
      Left            =   690
      TabIndex        =   3
      Top             =   2235
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
      Left            =   1275
      TabIndex        =   2
      Top             =   840
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Sigla 
      Height          =   315
      Left            =   5025
      TabIndex        =   1
      Top             =   300
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor6 
      Height          =   225
      Left            =   4080
      TabIndex        =   17
      Top             =   2580
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Valor5 
      Height          =   225
      Left            =   2850
      TabIndex        =   18
      Top             =   2580
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Valor4 
      Height          =   225
      Left            =   1605
      TabIndex        =   19
      Top             =   2625
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor8 
      Height          =   225
      Left            =   5625
      TabIndex        =   20
      Top             =   3480
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox VAlor7 
      Height          =   225
      Left            =   4395
      TabIndex        =   21
      Top             =   3480
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid GridCategoriaProduto 
      Height          =   2340
      Left            =   165
      TabIndex        =   8
      Top             =   1410
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   4128
      _Version        =   393216
      Rows            =   8
      Cols            =   6
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
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
      Left            =   285
      TabIndex        =   16
      Top             =   885
      Width           =   930
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
      Left            =   270
      TabIndex        =   15
      Top             =   375
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sigla:"
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
      Left            =   4470
      TabIndex        =   14
      Top             =   345
      Width           =   495
   End
End
Attribute VB_Name = "CategoriaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Alteração feita em 2/4/03 por Ivan
'inclusão de campos Valor 1, 2 e 3 e respectivos tratamentos

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGrid As AdmGrid

'variaveis do controle do grid
Dim iGrid_Item_Col As Integer
Dim iGrid_DescricaoItem_Col As Integer
Dim iGrid_Valor1_Col As Integer
Dim iGrid_Valor2_Col As Integer
Dim iGrid_Valor3_Col As Integer
Dim iGrid_Valor4_Col As Integer
Dim iGrid_Valor5_Col As Integer
Dim iGrid_Valor6_Col As Integer
Dim iGrid_Valor7_Col As Integer
Dim iGrid_Valor8_Col As Integer

Function Trata_Parametros(Optional objCategoriaProduto As ClassCategoriaProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma categoria foi passada por parametro
    If Not (objCategoriaProduto Is Nothing) Then

        lErro = Traz_CategoriaProduto_Tela(objCategoriaProduto)
        If lErro <> SUCESSO And lErro <> 19366 Then Error 22331
        
        'se a categoria nao está cadastrada
        If lErro = 19366 Then Categoria.Text = objCategoriaProduto.sCategoria

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 22331

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144287)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Categoria foi preenchida
    If Len(Categoria.Text) = 0 Then Error 22368
        
    'Preenche objCategoriaProduto
    objCategoriaProduto.sCategoria = Categoria.Text

    'Envia aviso perguntando se realmente deseja excluir Categoria
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CATEGORIAPRODUTO", objCategoriaProduto.sCategoria)

    If vbMsgRes = vbYes Then

        'Exclui Categoria
        lErro = CF("CategoriaProduto_Exclui", objCategoriaProduto)
        If lErro <> SUCESSO Then Error 22359

        'Exclui a Categoria da Combo
        For iIndice1 = 0 To Categoria.ListCount - 1

            If Categoria.List(iIndice1) = objCategoriaProduto.sCategoria Then
                Categoria.RemoveItem (iIndice1)
                Exit For
            End If

        Next
                
        Call Limpar_Tela
    
    End If
 
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 22359
        
        Case 22368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144288)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 22332

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case Err

        'Erros já tratados
        Case 22332

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144289)

    End Select

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 22333

    'Limpa a tela
    Call Limpar_Tela

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        'Erro já tratado
        Case 22333

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144290)

    End Select

    Exit Sub

End Sub

'Alterado por Ivan em 04/04/03
Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItensCategoria As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a Categoria está preenchida
    If Len(Trim(Categoria.Text)) = 0 Then gError 22334
    
    'Verifica se a Sigla está preenchida
    If Len(Trim(Sigla.Text)) = 0 Then gError 43172
    
    'Se não existir itens de categoria => erro
    If objGrid.iLinhasExistentes = 0 Then gError 116347

    'Chama Move_Tela_Memoria para passar os dados da tela para  os objetos
    lErro = Move_Tela_Memoria(objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO Then gError 22367

    lErro = Trata_Alteracao(objCategoriaProduto, objCategoriaProduto.sCategoria)
    If lErro <> SUCESSO Then gError 32315

    'Chama a função de gravacao
    lErro = CF("CategoriaProduto_Grava", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO Then gError 22351

    'Exclui ( se existir) da lista de Categoria
    Call ListaCategoria_Exclui(objCategoriaProduto.sCategoria)
    
    Categoria.AddItem objCategoriaProduto.sCategoria
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 22351, 22367, 32315

        Case 22334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", gErr)
            Categoria.SetFocus

        Case 43172
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_SIGLA_NAO_INFORMADA", gErr)
            Sigla.SetFocus

        Case 116347
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_SEM_ITEM_CORRESPONDENTE", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144291)

    End Select
    
    Exit Function

End Function

Private Sub ListaCategoria_Exclui(sCategoria As String)

Dim iIndice As Integer

    For iIndice = 0 To Categoria.ListCount - 1

        If Categoria.List(iIndice) = sCategoria Then

            Categoria.RemoveItem (iIndice)
            Exit For

        End If

    Next

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 22438

    'Limpa a Tela
    Call Limpar_Tela

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 22438

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144292)

    End Select

End Sub

Private Sub Limpar_Tela()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa os textos
    Categoria.Text = ""
    Sigla.Text = ""

    'Limpa GridCategoriaProduto
    Call Grid_Limpa(objGrid)

    'Linhas visíveis do grid
    objGrid.iLinhasExistentes = 0

    Categoria.SetFocus

    iAlterado = 0

End Sub

Private Sub Categoria_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Categoria_Click()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Categoria_Click

     iAlterado = REGISTRO_ALTERADO
    
     If Categoria.ListIndex <> -1 Then
        
        objCategoriaProduto.sCategoria = Categoria.Text
        
        lErro = Traz_CategoriaProduto_Tela(objCategoriaProduto)
        If lErro <> SUCESSO Then Error 22343

     End If

    Exit Sub

Erro_Categoria_Click:

    Select Case Err

        Case 22343

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144293)

    End Select

End Sub

Private Sub Categoria_Validate(bMantemFoco As Boolean)

Dim iIndice As Integer, lErro As Long
Dim colCategorias As New Collection

On Error GoTo Erro_Categoria_Validate

    If Len(Trim(Categoria.Text)) <> 0 Then
    
        If Categoria.ListIndex = -1 Then

            If Len(Trim(Categoria.Text)) > STRING_CATEGORIAPRODUTO_CATEGORIA Then Error 19365
            
            Call Combo_Item_Igual(Categoria)
                    
        End If
        
    End If

    Exit Sub
    
Erro_Categoria_Validate:

    Select Case Err

        Case 19365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_TAMMAX", Err, STRING_CATEGORIAPRODUTO_CATEGORIA)
'            Categoria.SetFocus
            bMantemFoco = True
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144294)

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
'Carrega a combo de categorias apenas com os códigos, sem a descrição

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega a ComboBox Categoria com os Códigos
    lErro = Carrega_Categoria()
    If lErro <> SUCESSO Then Error 22335

    'Inicializa o Grid
    Set objGrid = New AdmGrid
    
    lErro = Inicializa_Grid_CategoriaProduto(objGrid)
    If lErro <> SUCESSO Then Error 22336

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22335, 22336

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144295)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_Categoria() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaProduto As ClassCategoriaProduto

On Error GoTo Erro_Carrega_Categoria

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 22337

    For Each objCategoriaProduto In colCategorias

        'Insere na combo Categoria
        Categoria.AddItem objCategoriaProduto.sCategoria

    Next

    Carrega_Categoria = SUCESSO

    Exit Function

Erro_Carrega_Categoria:

    Carrega_Categoria = Err

    Select Case Err

        Case 22337

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144296)

    End Select

    Exit Function

End Function

'Alterado por Ivan
Private Function Inicializa_Grid_CategoriaProduto(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Valor 1")
    objGridInt.colColuna.Add ("Valor 2")
    objGridInt.colColuna.Add ("Valor 3")
    objGridInt.colColuna.Add ("Valor 4")
    objGridInt.colColuna.Add ("Valor 5")
    objGridInt.colColuna.Add ("Valor 6")
    objGridInt.colColuna.Add ("Valor 7")
    objGridInt.colColuna.Add ("Valor 8")

   'campos de edição do grid
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (Valor1.Name)
    objGridInt.colCampo.Add (Valor2.Name)
    objGridInt.colCampo.Add (Valor3.Name)
    objGridInt.colCampo.Add (Valor4.Name)
    objGridInt.colCampo.Add (Valor5.Name)
    objGridInt.colCampo.Add (Valor6.Name)
    objGridInt.colCampo.Add (VAlor7.Name)
    objGridInt.colCampo.Add (Valor8.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Item_Col = 1
    iGrid_DescricaoItem_Col = 2
    iGrid_Valor1_Col = 3
    iGrid_Valor2_Col = 4
    iGrid_Valor3_Col = 5
    iGrid_Valor4_Col = 6
    iGrid_Valor5_Col = 7
    iGrid_Valor6_Col = 8
    iGrid_Valor7_Col = 9
    iGrid_Valor8_Col = 10

    objGridInt.objGrid = GridCategoriaProduto

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_CATEGORIA + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 8

    'largura da 1ª coluna
    GridCategoriaProduto.ColWidth(0) = 400

    'largura automatica das demias colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoIncluirNoMeioGrid = 0

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_CategoriaProduto = SUCESSO

End Function

Sub GridCategoriaProduto_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Sub GridCategoriaProduto_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Sub GridCategoriaProduto_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Sub GridCategoriaProduto_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CategoriaProduto"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Categoria", Categoria.Text, STRING_CATEGORIAPRODUTO_CATEGORIA, "Categoria"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 22352

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144297)

    End Select

    Exit Sub

End Sub

'Alterado por Ivan
Function Move_Tela_Memoria(objCategoriaProduto As ClassCategoriaProduto, colItensCategoria As Collection) As Long
'Move os dados da tela para objCategoriaProduto

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objCategoriaProduto
    If Len(Trim(Categoria.Text)) > 0 Then objCategoriaProduto.sCategoria = Trim(Categoria.Text)
    If Len(Trim(Descricao.Text)) > 0 Then objCategoriaProduto.sDescricao = Descricao.Text
    If Len(Trim(Sigla.Text)) > 0 Then objCategoriaProduto.sSigla = Sigla.Text

    'Ir preenchendo uma colecao com todas as linhas "existentes" do grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Se o Item não estiver preenchido => erro
        If Len(Trim(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Item_Col))) = 0 Then Error 22353
        
        Set objCategoriaProdutoItem = New ClassCategoriaProdutoItem
        
        objCategoriaProdutoItem.sCategoria = objCategoriaProduto.sCategoria
        objCategoriaProdutoItem.sItem = Trim(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Item_Col))
        objCategoriaProdutoItem.iOrdem = iIndice
        objCategoriaProdutoItem.sDescricao = GridCategoriaProduto.TextMatrix(iIndice, iGrid_DescricaoItem_Col)
        objCategoriaProdutoItem.dvalor1 = StrParaDbl(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor1_Col))
        objCategoriaProdutoItem.dvalor2 = StrParaDbl(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor2_Col))
        objCategoriaProdutoItem.dvalor3 = StrParaDbl(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor3_Col))
        objCategoriaProdutoItem.dvalor4 = StrParaDbl(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor4_Col))
        objCategoriaProdutoItem.dvalor5 = StrParaDbl(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor5_Col))
        objCategoriaProdutoItem.dvalor6 = StrParaDbl(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor6_Col))
        objCategoriaProdutoItem.dvalor7 = StrParaDbl(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor7_Col))
        objCategoriaProdutoItem.dvalor8 = StrParaDbl(GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor8_Col))
    
        'Verifica se já existe o Item na coleção
        For iIndice1 = 1 To colItensCategoria.Count
            If UCase(objCategoriaProdutoItem.sItem) = UCase(colItensCategoria.Item(iIndice1).sItem) Then Error 22465
        Next
    
        colItensCategoria.Add objCategoriaProdutoItem

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err

        Case 22353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_ITEM_CATEGORIAPRODUTO2", Err, iIndice)

        Case 22439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", Err)

        Case 22465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_REPETIDO_NO_GRID", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144298)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Tela_Preenche

    objCategoriaProduto.sCategoria = colCampoValor.Item("Categoria").vValor

    If Len(objCategoriaProduto.sCategoria) > 0 Then

       'Traz dados da Categoria para a Tela
        lErro = Traz_CategoriaProduto_Tela(objCategoriaProduto)
        If lErro <> SUCESSO Then Error 22350

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 22350

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144299)

    End Select

    Exit Sub

End Sub

Function Traz_CategoriaProduto_Tela(objCategoriaProduto As ClassCategoriaProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItensCategoria As New Collection

On Error GoTo Erro_Traz_CategoriaProduto_Tela

    'Lê a tabela CategoriaProduto a partir da Categoria
    lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22540 Then Error 22414
    
    If lErro = 22540 Then Error 19366
    
    'Exibe os dados de objCategoriaProduto na tela
    Categoria.Text = objCategoriaProduto.sCategoria
    Descricao.Text = objCategoriaProduto.sDescricao
    Sigla.Text = objCategoriaProduto.sSigla

    'Lê a tabela CategoriaProdutoitem a partir da Categoria
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO And lErro <> 22541 Then Error 22358

    'Limpa o Grid antes de colocar algo nele
    Call Grid_Limpa(objGrid)

    'Exibe os dados da coleção na tela
    For iIndice = 1 To colItensCategoria.Count
        
        'Insere no Grid CategoriaProduto
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = colItensCategoria.Item(iIndice).sDescricao
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Item_Col) = colItensCategoria.Item(iIndice).sItem
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor1_Col) = Formata_ValorCateg(colItensCategoria.Item(iIndice).dvalor1)
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor2_Col) = Formata_ValorCateg(colItensCategoria.Item(iIndice).dvalor2)
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor3_Col) = Formata_ValorCateg(colItensCategoria.Item(iIndice).dvalor3)
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor4_Col) = Formata_ValorCateg(colItensCategoria.Item(iIndice).dvalor4)
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor5_Col) = Formata_ValorCateg(colItensCategoria.Item(iIndice).dvalor5)
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor6_Col) = Formata_ValorCateg(colItensCategoria.Item(iIndice).dvalor6)
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor7_Col) = Formata_ValorCateg(colItensCategoria.Item(iIndice).dvalor7)
        GridCategoriaProduto.TextMatrix(iIndice, iGrid_Valor8_Col) = Formata_ValorCateg(colItensCategoria.Item(iIndice).dvalor8)

    Next

    objGrid.iLinhasExistentes = colItensCategoria.Count

    iAlterado = 0

    Traz_CategoriaProduto_Tela = SUCESSO

    Exit Function

Erro_Traz_CategoriaProduto_Tela:

    Traz_CategoriaProduto_Tela = Err

    Select Case Err

        Case 22358, 22414, 19366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144300)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

''''    Me.ValidateControls

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGrid = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
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

        Select Case GridCategoriaProduto.Col

            Case iGrid_Item_Col

                lErro = Saida_Celula_Item(objGridInt)
                If lErro <> SUCESSO Then gError 22344

            Case iGrid_DescricaoItem_Col

                lErro = Saida_Celula_DescricaoItem(objGridInt)
                If lErro <> SUCESSO Then gError 22345

            Case iGrid_Valor1_Col
                
                'Faz a saída do controle passado como parâmetro
                lErro = Saida_Celula_Valor(objGridInt, Valor1)
                If lErro <> SUCESSO Then gError 116343

            Case iGrid_Valor2_Col
                
                'Faz a saída do controle passado como parâmetro
                lErro = Saida_Celula_Valor(objGridInt, Valor2)
                If lErro <> SUCESSO Then gError 116344
            
            Case iGrid_Valor3_Col
                
                'Faz a saída do controle passado como parâmetro
                lErro = Saida_Celula_Valor(objGridInt, Valor3)
                If lErro <> SUCESSO Then gError 116345
                
            Case iGrid_Valor4_Col
                
                'Faz a saída do controle passado como parâmetro
                lErro = Saida_Celula_Valor(objGridInt, Valor4)
                If lErro <> SUCESSO Then gError 116343

            Case iGrid_Valor5_Col
                
                'Faz a saída do controle passado como parâmetro
                lErro = Saida_Celula_Valor(objGridInt, Valor5)
                If lErro <> SUCESSO Then gError 116343

            Case iGrid_Valor6_Col
                
                'Faz a saída do controle passado como parâmetro
                lErro = Saida_Celula_Valor(objGridInt, Valor6)
                If lErro <> SUCESSO Then gError 116343

            Case iGrid_Valor7_Col
                
                'Faz a saída do controle passado como parâmetro
                lErro = Saida_Celula_Valor(objGridInt, VAlor7)
                If lErro <> SUCESSO Then gError 116343

            Case iGrid_Valor8_Col
                
                'Faz a saída do controle passado como parâmetro
                lErro = Saida_Celula_Valor(objGridInt, Valor8)
                If lErro <> SUCESSO Then gError 116343

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 22346

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 22346
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 22344, 22345, 116343, 116344, 116345

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144301)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    'Se o campo foi preenchido
    If Len(Trim(DescricaoItem.Text)) > 0 Then
        
        'verifica se precisa preencher uma o grid com uma nova linha
        If GridCategoriaProduto.Row - GridCategoriaProduto.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144302)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Item(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Item

    Set objGridInt.objControle = Item

    'Se o campo foi preenchido
    If Len(Trim(Item.Text)) > 0 Then
        'verifica se precisa preencher uma o grid com uma nova linha
        If GridCategoriaProduto.Row - GridCategoriaProduto.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22347

    Saida_Celula_Item = SUCESSO

    Exit Function

Erro_Saida_Celula_Item:

    Saida_Celula_Item = Err

    Select Case Err

        Case 22347
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144303)

    End Select

    Exit Function

End Function

'alteração de saida //Ivan
Private Function Saida_Celula_Valor(objGridInt As AdmGrid, objControle As Object) As Long
'Faz a crítica da célula Valor (1,2,3//recebida como parametro) do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sValor As String

On Error GoTo Erro_Saida_Celula_Valor

    'seta o obj c/ o controle
    Set objGridInt.objControle = objControle

    'Se o controle estiver preenchido
    If Len(Trim(objControle.Text)) > 0 Then
        
        'verifica se o valor é valido
        lErro = Valor_Double_Critica(objGridInt.objControle)
        If lErro <> SUCESSO Then gError 116342
        
        objControle.Text = Formata_ValorCateg(StrParaDbl(objControle.Text))
    
        'verifica se precisa preencher uma o grid com uma nova linha
        If GridCategoriaProduto.Row - GridCategoriaProduto.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    'abandona a célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 116346

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 116346
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 116342
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO", gErr, objControle.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144304)

    End Select

    Exit Function

End Function

Private Sub GridCategoriaProduto_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridCategoriaProduto_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridCategoriaProduto_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Private Sub GridCategoriaProduto_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridCategoriaProduto_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Private Sub Item_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sigla_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

'alterado por Ivan 2/4/03
Private Sub Valor1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor4_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor5_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor6_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor7_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor8_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor2_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor2_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor3_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor3_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor4_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor4_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor5_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor5_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor6_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor6_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor7_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor7_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor8_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor8_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor1
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor2
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor3
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor4_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor4
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor5_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor5
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor6_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor6
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor7_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = VAlor7
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor8_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor8
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CATEGORIA_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Categoria de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CategoriaProduto"
    
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

'**** fim do trecho a ser copiado *****


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

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Function Formata_ValorCateg(ByVal dValor As Double) As String
Dim sResult As String

    sResult = Format(dValor, "##,###,###.#####")
    If Right(sResult, 1) = "," Then sResult = Mid(sResult, 1, Len(sResult) - 1)
    
    Formata_ValorCateg = sResult
    
End Function
