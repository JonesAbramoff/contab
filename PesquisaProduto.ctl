VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PesquisaProduto 
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   ScaleHeight     =   2955
   ScaleWidth      =   6795
   Begin VB.ComboBox TipoProduto 
      Height          =   315
      ItemData        =   "PesquisaProduto.ctx":0000
      Left            =   1245
      List            =   "PesquisaProduto.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   3165
   End
   Begin VB.CommandButton BotaoPesquisa 
      Caption         =   "Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4875
      TabIndex        =   6
      Top             =   855
      Width           =   1785
   End
   Begin VB.Frame Frame7 
      Caption         =   "Categorias"
      Height          =   1770
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   1050
      Width           =   4320
      Begin VB.ComboBox CategoriaProdutoItem 
         Height          =   315
         Left            =   2310
         TabIndex        =   5
         Top             =   375
         Width           =   1632
      End
      Begin VB.ComboBox CategoriaProduto 
         Height          =   315
         Left            =   735
         TabIndex        =   4
         Top             =   375
         Width           =   1548
      End
      Begin MSFlexGridLib.MSFlexGrid GridCategoria 
         Height          =   1320
         Left            =   270
         TabIndex        =   3
         Top             =   285
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   2328
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5400
      ScaleHeight     =   495
      ScaleWidth      =   1185
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1245
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   105
         Picture         =   "PesquisaProduto.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   615
         Picture         =   "PesquisaProduto.ctx":0536
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   615
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label LblFornecedor 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
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
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   645
      Width           =   1035
   End
   Begin VB.Label LblTipoProduto 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Left            =   705
      TabIndex        =   9
      Top             =   195
      Width           =   450
   End
End
Attribute VB_Name = "PesquisaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''Verificar a necessidade de passar um objProduto no Trata_Param)

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjEventoProduto As AdmEvento

Dim iTipoProdutoAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iAlterado As Integer

Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

 'Chama Limpa_Tela
  Call LimpaTela_PesquisaProduto

End Sub

Private Sub BotaoPesquisa_Click()
'Chama browse com Produtos que satisfazem filtros acompanhados de preços e quant disponíveis em Estoque
'e fecha essa tela.
'Ao selecionar no browse volta para a tela que chamadora da tela de pesquisa se houver e lá coloca o Produto
'com os dados pertinentes.

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection
Dim sSelecao As String
Dim sSelecaoCategoria As String
Dim iPreenchido As Integer
Dim objProdutoCategoria As ClassProdutoCategoria

On Error GoTo Erro_Botao_Pesquisa_Click

    'Carrega os dados da tela em objProduto e objFornecedor
    lErro = Move_Tela_Memoria(objProduto, objFornecedor)
    If lErro <> SUCESSO Then gError 80150
            
    '??? ok
    'Tratamento de browse dinâmico
    'Verifica o preenchimento dos campos antes deles serem incluídos na coleção
    'que deverá ser passada como parâmetro no chama_tela
    'Prepara a passagem de parâmetro do browse dinâmico em sSelecao
    
    'Verifica se o tipo do produto está preenchido
    If objProduto.iTipo <> 0 Then
        sSelecao = "Tipo = ?"
        iPreenchido = 1
        'Se estiver --> adiciona na coleção
        colSelecao.Add (objProduto.iTipo)
    End If
    
    'Verifica se o nome reduzido do fornecedor está preenchido
    If Len(Trim(objFornecedor.sNomeReduzido)) > 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Fornecedor = ?"
        Else
            iPreenchido = 1
            sSelecao = "Fornecedor = ?"
        End If
        'Se estiver --> adiciona na coleção
        colSelecao.Add (objFornecedor.sNomeReduzido)
    End If
        
    iPreenchido = 0
    
    For Each objProdutoCategoria In objProduto.colCategoriaItem
        
        'Prepara a passagem de parâmetro do browse dinâmico em sSelecaoCategoria
        'Verifica se a categoria do produto está preenchida
        If Len(Trim(objProdutoCategoria.sCategoria)) > 0 Then
            'Se Categoria estiver preenhida
            If iPreenchido = 1 Then
                'Adicionar categoria como parametro no Select
                sSelecaoCategoria = sSelecaoCategoria & " OR (Categoria = ? "
                'adiciona na coleção
                colSelecao.Add (objProdutoCategoria.sCategoria)
            Else
                sSelecaoCategoria = " ((Categoria = ? "
                'adiciona na coleção
                colSelecao.Add (objProdutoCategoria.sCategoria)
                iPreenchido = 1
            End If
            
            'Verifica se o item da categoria está preenchida
            If Len(Trim(objProdutoCategoria.sItem)) > 0 Then
                'Adicionar o item como parametro no Select
                sSelecaoCategoria = sSelecaoCategoria & " AND Item = ? "
                'adiciona na coleção
                colSelecao.Add (objProdutoCategoria.sItem)
            End If
            'Finaliza a passagem de parametros no Select
            sSelecaoCategoria = sSelecaoCategoria & ") "
        End If
        
    Next
    
    'carrrega sSelecao com os dados de sSelecao e sSelecaoCategoria
    If Len(Trim(sSelecaoCategoria)) > 0 Then
        If Len(Trim(sSelecao)) > 0 Then
            sSelecao = sSelecao & " AND " & sSelecaoCategoria & ")"
        Else
            sSelecao = sSelecaoCategoria & ")"
        End If
    End If
        
    'Chama a tela de browse de acordo com os filtros passados
    Call Chama_Tela("ProdutoVendaLojaLista", colSelecao, objProduto, gobjEventoProduto, sSelecao)
        
    Unload Me
    
    Exit Sub
    
Erro_Botao_Pesquisa_Click:

    Select Case gErr
                
        Case 80150
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164835)
    
    End Select
    
    Exit Sub
    
End Sub
        
Function Trata_Parametros(Optional objEventoProduto As AdmEvento) As Long

    Set gobjEventoProduto = objEventoProduto
    
    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega a ComboBox de produtos
    lErro = Carrega_TipoProduto()
    If lErro <> SUCESSO Then gError 80130
    
    'Carrega a ComboBox de Categorias
    lErro = Carrega_CategoriaProduto()
    If lErro <> SUCESSO Then gError 80132
       
    'Inicialiaza o Grid de Categoria
    Set objGridCategoria = New AdmGrid
    
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError 80134 '??gerror ok

    '???? IDENTAÇÃO ok
    lErro_Chama_Tela = SUCESSO
                
    iAlterado = 0
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 80130, 80132, 80134

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164836)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    '????? FALTA CÓDIGO ok
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    '????? FALTA CÓDIGO ok
    'Libera a variável global
    Set gobjEventoProduto = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Pesquisa Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PesquisaProduto"
    
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

Private Sub Fornecedor_GotFocus()

    iFornecedorAlterado = REGISTRO_ALTERADO

End Sub


Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedorProduto As New ClassFornecedor
Dim iCodFilial As Integer
    
On Error GoTo Erro_Fornecedor_Validate
    
    'Verifica se fornecedor foi alterado
    If iFornecedorAlterado <> REGISTRO_ALTERADO Then Exit Sub
        
    '???? É ASSIM QUE VOCÊ VERIFICA SE O FORNECEDOR FOI INFORMADO NA TELA? ok
    'Verifica se foi informado fornecedor
    If Len(Trim(Fornecedor.Text)) = 0 Then Exit Sub
    
    'Verifica se fornecedor esta cadastrados no BD
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedorProduto, iCodFilial, False)
    If lErro <> SUCESSO And lErro <> 6663 And lErro <> 6664 Then gError 80140
    If lErro <> SUCESSO Then gError 80148
       
    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 80140
        
        Case 80148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164837)
    
    End Select

    Exit Sub

End Sub

Private Sub TipoProduto_Change()

    iTipoProdutoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoProduto_Click()

    iTipoProdutoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoProduto_GotFocus()

    iTipoProdutoAlterado = 0

End Sub

Private Sub TipoProduto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_TipoProduto_Validate

    If iTipoProdutoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TipoProduto.Text)) = 0 Then Exit Sub

    '?????? PORQUE VC ESTÁ VERIFICANDO SE O ÚLTIMO TIPOPRODUTO
    '???? INCLUÍDO ESTÁ EM UMA POSIÇÃO <> 0? ok
    If TipoProduto.ListIndex = -1 Then Exit Sub
        
    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TipoProduto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 80137

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoProduto.iTipo = iCodigo

        'Tenta ler TipoDeProduto com esse código no BD
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 80138
        If lErro <> SUCESSO Then gError 80138 'Não encontrou Tabela Preço no BD

        'Encontrou TipoDeProduto no BD, coloca no Text da Combo
        TipoProduto.Text = CStr(objTipoProduto.iTipo) & SEPARADOR & objTipoProduto.sDescricao

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 80139

    Exit Sub

Erro_TipoProduto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 80137
    
    
        Case 80138, 80139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164838)

    End Select

    Exit Sub

End Sub
    
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


Function Carrega_TipoProduto() As Long
'Carrega na combo os tipos de produtos existentes

Dim lErro As Long
Dim objCodDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TipoProduto

    'Lê o código e a descrição de todas as Tabelas de Preço
    lErro = CF("Cod_Nomes_Le", "TiposDeProduto", "TipoDeProduto", "Descricao", STRING_TIPODEPRODUTO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 80131
    
    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na Lista de Tabela de Produtos
        TipoProduto.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoProduto.ItemData(TipoProduto.NewIndex) = objCodDescricao.iCodigo
    
    Next

    Carrega_TipoProduto = SUCESSO

    Exit Function

Erro_Carrega_TipoProduto:

    Carrega_TipoProduto = gErr

    Select Case gErr

        Case 80131

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164839)

    End Select

    Exit Function

End Function

Private Function Carrega_CategoriaProduto() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Carrega_CategoriaProduto

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 80133

    For Each objCategoriaProduto In colCategorias
        'Insere na combo CategoriaProduto
        CategoriaProduto.AddItem objCategoriaProduto.sCategoria
    Next

    Carrega_CategoriaProduto = SUCESSO

    Exit Function

Erro_Carrega_CategoriaProduto:

    Carrega_CategoriaProduto = gErr

    Select Case gErr

        Case 80133

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164840)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Categoria

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Item")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (CategoriaProduto.Name)
    objGridInt.colCampo.Add (CategoriaProdutoItem.Name)

    'Colunas do Grid
    iGrid_Categoria_Col = 1
    iGrid_Valor_Col = 2
    
    'Grid do GridInterno
    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridCategoria.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)
'Rotina que habilita a entrada na celula

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Rotina_Grid_Enable

    If objControl.Name = CategoriaProdutoItem.Name Then
        
        If Len(Trim(GridCategoria.TextMatrix(iLinha, iGrid_Categoria_Col))) = 0 Then
            objControl.Enabled = False

        Else
            objControl.Enabled = True
            
            objCategoriaProduto.sCategoria = GridCategoria.TextMatrix(iLinha, iGrid_Categoria_Col)
            
            lErro = Carrega_ComboCategoriaProdutoItem(objCategoriaProduto)
            If lErro <> SUCESSO Then gError 80135
        
        End If

    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 80135
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164841)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ComboCategoriaProdutoItem(objCategoriaProduto As ClassCategoriaProduto) As Long
'Carrega o Item da Categoria na Combobox

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim sComboItem As String

On Error GoTo Erro_Carrega_ComboCategoriaProdutoItem

    sComboItem = CategoriaProdutoItem.Text

    CategoriaProdutoItem.Clear
         
   'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO Then gError 80136

    'Insere na combo CategoriaProdutoItem
    For Each objCategoriaProdutoItem In colItensCategoria
        'Insere na combo CategoriaProduto
        CategoriaProdutoItem.AddItem objCategoriaProdutoItem.sItem
    Next
    
    CategoriaProdutoItem.Text = sComboItem

    Carrega_ComboCategoriaProdutoItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaProdutoItem:

    Carrega_ComboCategoriaProdutoItem = gErr

    Select Case gErr

        Case 80136

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164842)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'CategoriaProduto
            Case iGrid_Categoria_Col
                lErro = Saida_Celula_Categoria(objGridInt)
                If lErro <> SUCESSO Then gError 80142
            
            'CategoriaProdutoItem
            Case iGrid_Valor_Col
                lErro = Saida_Celula_Item(objGridInt)
                If lErro <> SUCESSO Then gError 80143
                        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 80144

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 80142, 80143
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 80144

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164843)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = CategoriaProduto

    'Se necessário cria uma nova linha no Grid
    If Len(Trim(CategoriaProduto.Text)) > 0 Then
    
        If GridCategoria.Row > objGridInt.iLinhasExistentes Then
            objGridCategoria.iLinhasExistentes = objGridCategoria.iLinhasExistentes + 1
        End If
    
    Else
        GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Valor_Col) = ""
                    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 80141

    Saida_Celula_Categoria = SUCESSO
    
    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = gErr

    Select Case gErr

        Case 80141
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164844)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Item(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Item

    Set objGridInt.objControle = CategoriaProdutoItem

    If Len(Trim(CategoriaProdutoItem.Text)) > 0 Then

        'Verifica se já existe a categoria no Grid
        For iIndice = 1 To objGridCategoria.iLinhasExistentes

            If iIndice <> GridCategoria.Row Then
        
                If GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = CategoriaProduto.Text And GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = CategoriaProdutoItem.Text Then gError 80146
           End If
        Next
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 80147

    '????? IDENTAÇãO ok
    Saida_Celula_Item = SUCESSO

    Exit Function

Erro_Saida_Celula_Item:

    Saida_Celula_Item = gErr

    Select Case gErr
    
        Case 80145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", gErr, Error$)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 80146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_EXISTE", gErr, CategoriaProdutoItem, CategoriaProduto)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 80147
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164845)
    
    End Select
            
    Exit Function
    
End Function


Private Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub


Private Sub GridCategoria_EnterCell()
    
    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)

End Sub

Private Sub GridCategoria_GotFocus()

    Call Grid_Recebe_Foco(objGridCategoria)

End Sub
Private Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)

End Sub
Private Sub GridCategoria_LostFocus()

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub
Private Sub GridCategoria_LeaveCell()

    Call Saida_Celula(objGridCategoria)

End Sub
Private Sub CategoriaProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Private Sub CategoriaProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = CategoriaProduto
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub CategoriaProdutoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Private Sub CategoriaProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Private Sub CategoriaProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = CategoriaProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub GridCategoria_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Private Sub GridCategoria_RowColChange()

    Call Grid_RowColChange(objGridCategoria)

End Sub

Private Sub GridCategoria_Scroll()

    Call Grid_Scroll(objGridCategoria)

End Sub

Private Function LimpaTela_PesquisaProduto()

    Call Limpa_Tela(Me)
  
    Call Grid_Limpa(objGridCategoria)
  
    Fornecedor.Text = ""
    TipoProduto.Text = ""
    
End Function


Function Move_Tela_Memoria(objProduto As ClassProduto, objFornecedor As ClassFornecedor)
'Move os dados da memoria para os obj(s)
Dim iIndice As Integer
Dim lErro As Long
Dim objProdutoCategoria As New ClassProdutoCategoria

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se os campos estao preenchidos
    If Len(Trim(TipoProduto.Text)) > 0 Then
        objProduto.iTipo = Codigo_Extrai(TipoProduto.Text)
    End If
         
    If Len(Trim(Fornecedor.Text)) > 0 Then
                    
        objFornecedor.sNomeReduzido = Fornecedor.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6722 Then gError 80149
        If lErro = 6722 Then gError 80151
        
    End If
     
    'Retira os dados do Grid e guarda no objProdutoCategoria
    For iIndice = 1 To objGridCategoria.iLinhasExistentes

        Set objProdutoCategoria = New ClassProdutoCategoria

        objProdutoCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col)
        objProdutoCategoria.sItem = GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col)

        objProduto.colCategoriaItem.Add objProdutoCategoria

    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 80149
    
        Case 80151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, Fornecedor)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164846)

    End Select
    
    Exit Function
    
End Function
