VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl CorTonTPint 
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ScaleHeight     =   5895
   ScaleWidth      =   5040
   Begin VB.CommandButton BotaoAutomatico 
      Caption         =   "Cadastrar todas as combinações de cores, modelagens e tipos de pintura"
      Height          =   585
      Left            =   90
      TabIndex        =   10
      Top             =   150
      Width           =   3030
   End
   Begin VB.ComboBox Almoxarifado 
      Height          =   315
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   945
      Width           =   3405
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   3405
      TabIndex        =   7
      Top             =   1890
      Width           =   1155
      _ExtentX        =   2037
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
      Format          =   "#,##0.00###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   225
      Left            =   2370
      TabIndex        =   6
      Top             =   1890
      Width           =   990
      _ExtentX        =   1746
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
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3225
      ScaleHeight     =   495
      ScaleWidth      =   1575
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   150
      Width           =   1635
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "CorTonTPint.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "CorTonTPint.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CorTonTPint.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Item 
      Height          =   225
      Left            =   585
      TabIndex        =   1
      Top             =   1890
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   7
      Mask            =   "9999999"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridItens 
      Height          =   4455
      Left            =   135
      TabIndex        =   0
      Top             =   1335
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   8
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
   End
   Begin VB.Label AlmoxarifadoLabel 
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1005
      Width           =   1155
   End
End
Attribute VB_Name = "CorTonTPint"
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
Dim gobjProduto As ClassProduto

'variaveis do controle do grid
Dim iGrid_Item_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Valor_Col As Integer

Function Trata_Parametros(objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjProduto = objProduto
    
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 178501
    
    If lErro <> SUCESSO Then gError 178502
    
    If objProduto.iGerencial = NAO_GERENCIAL Then gError 178503
    
    If Right(objProduto.sCodigo, 7) <> "0000000" Then gError 178504
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 178501

        Case 178502
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 178503
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 178504
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_TER_CORTONPINT", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178467)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoAutomatico_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iAlmoxarifado As Integer
Dim colItens As New Collection

On Error GoTo Erro_BotaoAutomatico_Click

    If Almoxarifado.ListIndex = -1 Or Len(Trim(Almoxarifado.Text)) = 0 Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_NAO_SELECIONADO")

            If vbMsgRes = vbNo Then gError 178526
    
    Else

            iAlmoxarifado = Almoxarifado.ItemData(Almoxarifado.ListIndex)
    End If


    GL_objMDIForm.MousePointer = vbHourglass

    lErro = CF("Produto_Grava_Harmonia", giFilialEmpresa, gobjProduto, colItens, iAlmoxarifado, 1)
    If lErro <> SUCESSO Then gError 178527

    GL_objMDIForm.MousePointer = vbDefault

    iAlterado = 0
    
    Unload Me
    
    Exit Sub

Erro_BotaoAutomatico_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 178526, 178527
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178528)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 178468

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 178468

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178469)

    End Select

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 178470

    Unload Me

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 178470

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178471)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim colItens As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim iAlmoxarifado As Integer

On Error GoTo Erro_Gravar_Registro

    'Se não existir itens de categoria => erro
    If objGrid.iLinhasExistentes = 0 Then gError 178472

    If Almoxarifado.ListIndex = -1 Or Len(Trim(Almoxarifado.Text)) = 0 Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_NAO_SELECIONADO")

            If vbMsgRes = vbNo Then gError 178522
    
    Else

            iAlmoxarifado = Almoxarifado.ItemData(Almoxarifado.ListIndex)
    End If

    'Chama Move_Tela_Memoria para passar os dados da tela para  os objetos
    lErro = Move_Tela_Memoria(colItens)
    If lErro <> SUCESSO Then gError 178473
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("Produto_Grava_Harmonia", giFilialEmpresa, gobjProduto, colItens, iAlmoxarifado)
    If lErro <> SUCESSO Then gError 178498
    
    GL_objMDIForm.MousePointer = vbDefault
    
    iAlterado = 0
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr

    Select Case gErr

        Case 178472
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOS_NAO_INFORMADOS", gErr)
        
        Case 178473, 178498, 178522
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178474)

    End Select
    
    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 178475

    'Limpa a Tela
    Call Limpar_Tela

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 178475

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178476)

    End Select

End Sub

Private Sub Limpar_Tela()

Dim lErro As Long

    Call Grid_Limpa(objGrid)

    iAlterado = 0

End Sub

Public Sub Form_Load()
'Carrega a combo de categorias apenas com os códigos, sem a descrição

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objGrid = New AdmGrid

    lErro = Carga_Almoxarifado()
    If lErro <> SUCESSO Then gError 178521

    lErro = Inicializa_Grid_Itens(objGrid)
    If lErro <> SUCESSO Then gError 178477

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 178477, 178521

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178478)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cor Tonalidade Pintura")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Valor")

   'campos de edição do grid
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Valor.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Item_Col = 1
    iGrid_Quantidade_Col = 2
    iGrid_Valor_Col = 3

    objGridInt.objGrid = GridItens

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 1000

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 14

    'largura da 1ª coluna
    GridItens.ColWidth(0) = 400

    'largura automatica das demias colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

End Function

Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Sub GridItens_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Sub GridItens_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Function Move_Tela_Memoria(colItens As Collection) As Long
'Move os dados da tela para objCategoriaProduto

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim objEstoqueProduto As ClassEstoqueProduto

On Error GoTo Erro_Move_Tela_Memoria

    'Ir preenchendo uma colecao com todas as linhas "existentes" do grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        Set objEstoqueProduto = New ClassEstoqueProduto

        objEstoqueProduto.sProduto = GridItens.TextMatrix(iIndice, iGrid_Item_Col)
        objEstoqueProduto.dQuantidadeInicial = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objEstoqueProduto.dSaldoInicial = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Valor_Col))

        colItens.Add objEstoqueProduto
                 
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178478)

    End Select

    Exit Function

End Function


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGrid = Nothing

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

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridItens.Col

            Case iGrid_Item_Col

                lErro = Saida_Celula_Item(objGridInt)
                If lErro <> SUCESSO Then gError 178479

            Case iGrid_Quantidade_Col

                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 178511

            Case iGrid_Valor_Col

                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then gError 178512


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 178480

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 178479, 178511, 178512
        
        Case 178480
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178481)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Item(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim colItensCategoria As New Collection
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim iAchou As Integer

On Error GoTo Erro_Saida_Celula_Item

    Set objGridInt.objControle = Item

    'Se o campo foi preenchido
    If Len(Trim(Item.Text)) > 0 Then
        
        If Len(Trim(Item.Text)) < 7 Then gError 178488
        
        objCategoriaProduto.sCategoria = "Cor"
        
        'Le na tabela de CategoriaProdutoItem todos os itens de uma Categoria e os retorna na coleção colItensCategoria
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 178482
                 
        For Each objCategoriaProdutoItem In colItensCategoria
            If StrParaDbl(Left(Item.Text, 2)) = objCategoriaProdutoItem.dvalor1 Then
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 0 Then gError 178483
        
        iAchou = 0
        
        Set colItensCategoria = New Collection
        
        objCategoriaProduto.sCategoria = "Tonalidade"
        
        'Le na tabela de CategoriaProdutoItem todos os itens de uma Categoria e os retorna na coleção colItensCategoria
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 178484
                 
        For Each objCategoriaProdutoItem In colItensCategoria
            If StrParaDbl(Mid(Item.Text, 3, 3)) = objCategoriaProdutoItem.dvalor1 Then
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 0 Then gError 178485
        
        iAchou = 0
        
        Set colItensCategoria = New Collection
        
        objCategoriaProduto.sCategoria = "Tipo de Pintura"
        
        'Le na tabela de CategoriaProdutoItem todos os itens de uma Categoria e os retorna na coleção colItensCategoria
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 178486
                 
        For Each objCategoriaProdutoItem In colItensCategoria
            If StrParaDbl(Mid(Item.Text, 6, 2)) = objCategoriaProdutoItem.dvalor1 Then
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 0 Then gError 178487
        
        
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> GridItens.Row And GridItens.TextMatrix(iIndice, iGrid_Item_Col) = Item.Text Then gError 178489
        Next
        
        'verifica se precisa preencher uma o grid com uma nova linha
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 22347

    Saida_Celula_Item = SUCESSO

    Exit Function

Erro_Saida_Celula_Item:

    Saida_Celula_Item = gErr

    Select Case gErr

        Case 178483
            Call Rotina_Erro(vbOKOnly, "ERRO_COR_NAO_CADASTRADA", gErr, Left(Item.Text, 2))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178485
            Call Rotina_Erro(vbOKOnly, "ERRO_TONALIDADE_NAO_CADASTRADA", gErr, Mid(Item.Text, 3, 3))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 178487
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPINTURA_NAO_CADASTRADA", gErr, Mid(Item.Text, 6, 2))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178488
            Call Rotina_Erro(vbOKOnly, "ERRO_CORTONPINT_MENOR_7_DIGITOS", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178489
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_REPETIDO_NO_GRID", gErr, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 22347
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178490)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 178513

        dQuantidade = CDbl(Quantidade.Text)

        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178514

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 178513, 178514
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178515)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorUnitario As Double
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    'Se estiver preenchido
    If Len(Trim(Valor.ClipText)) > 0 Then
    
        'Faz a crítica do valor
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 178516

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 178517
    
    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 178516, 178517
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178518)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cor Tonalidade Tipo de Pintura"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CorTonTPint"
    
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

Function Carga_Almoxarifado() As Long

Dim lErro As Long
Dim objAlmoxarifado As ClassAlmoxarifado
Dim colAlmoxarifados As New Collection

On Error GoTo Erro_Carga_Almoxarifado

    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa", giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then gError 178519

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifado.AddItem objAlmoxarifado.sNomeReduzido
        Almoxarifado.ItemData(Almoxarifado.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Almoxarifado.AddItem ""

    Carga_Almoxarifado = SUCESSO
    
    Exit Function
    
Erro_Carga_Almoxarifado:

    Carga_Almoxarifado = gErr

    Select Case gErr
    
        Case 178519

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178520)

    End Select
    
    Exit Function

End Function

