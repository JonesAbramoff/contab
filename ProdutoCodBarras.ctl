VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl ProdutoCodBarras 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3795
   ScaleWidth      =   4800
   Begin VB.CommandButton BotaoCancel 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3760
      TabIndex        =   5
      Top             =   150
      Width           =   930
   End
   Begin VB.CommandButton BotaoRetornar 
      Caption         =   "Retornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   4
      Top             =   150
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Códigos"
      Height          =   3105
      Left            =   75
      TabIndex        =   2
      Top             =   600
      Width           =   4635
      Begin VB.TextBox CodBarras 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2355
      End
      Begin MSFlexGridLib.MSFlexGrid GridCodBarras 
         Height          =   2640
         Left            =   105
         TabIndex        =   3
         Top             =   300
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   4657
         _Version        =   393216
         Rows            =   6
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
      End
   End
   Begin VB.Label Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   975
      TabIndex        =   1
      Top             =   180
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Produto :"
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
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   0
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "ProdutoCodBarras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()


'***ALTERACAO POR TULIO EM 28/05***
'Constante que indica o numero de linhas visiveis no grid
'de codigo de barras
Const LINHAS_VISIVEIS_GRIDCODBARRAS = 6



'objProduto global a tela
Dim gobjProduto As ClassProduto

'Variavel que representa o grid
Public objGridCodBarras As AdmGrid

'Colunas do grid
Dim iGrid_CodBarras_Col As Integer
'***FIM ALTERACAO POR TULIO EM 28/05***

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'instancia o objGrid
    Set objGridCodBarras = New AdmGrid

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165569)
            
    End Select
    
    Exit Sub
    
End Sub

'***ALTERACAO POR TULIO EM 28/05***
Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sProdutoMascarado As String

On Error GoTo Erro_Trata_Parametros

    'instancia gobjproduto como o parametro objproduto
    Set gobjProduto = objProduto

    'se produto estiver preenchido
    If Len(Trim(gobjProduto.sCodigo)) > 0 Then

        'Mascarar produto
        lErro = Mascara_MascararProduto(gobjProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 101707
        
        'Coloca o Codigo na tela
        Produto.Caption = sProdutoMascarado
    
    Else
    
        'certifica-se q o label de produto nao estara preenchido
        Produto.Caption = STRING_VAZIO
    
    End If
    
    '***ALTERACAO POR TULIO EM 28/05***
    'Executa inicializacao do Grid
    lErro = Inicializa_GridCodBarras(objGridCodBarras)
    If lErro <> SUCESSO Then gError 101705
    '***FIM ALTERACAO POR TULIO EM 28/05***
    
    'preenche o grid com os codigos de barras passados
    lErro = Carrega_GridCodBarras()
    If lErro <> SUCESSO Then gError 101708

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
        
        Case 101705, 101707, 101708
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165570)
            
    End Select

    Exit Function

End Function
'***FIM ALTERACAO POR TULIO EM 28/05***

'***ALTERACAO POR TULIO EM 28/05***
Private Function Carrega_GridCodBarras() As Long
'carrega o grid de codigo de barras a partir da colecao de codigo de barras
'que esta no gobjProduto

Dim vCodigoBarra As Variant
Dim iLinha As Integer

On Error GoTo Erro_Carrega_GridCodBarras

    'inicializa a variavel como 1a linha do grid
    iLinha = 1

    'Para cada item na colecao
    For Each vCodigoBarra In gobjProduto.colCodBarras

        'Insere no Grid de codigo de barras
        GridCodBarras.TextMatrix(iLinha, iGrid_CodBarras_Col) = CStr(vCodigoBarra)
        
        'incrementa da linha
        iLinha = iLinha + 1
                
    Next

    'atualiza linhas existentes
    objGridCodBarras.iLinhasExistentes = gobjProduto.colCodBarras.Count
    
    Carrega_GridCodBarras = SUCESSO
    
    Exit Function

Erro_Carrega_GridCodBarras:

    Carrega_GridCodBarras = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165571)
            
    End Select
        
    Exit Function

End Function
'***FIM ALTERACAO POR TULIO EM 28/05***

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
''''    Parent.HelpContextID = IDH_KIT
    Set Form_Load_Ocx = Me
    Caption = "Códigos de Barras"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ProdutoCodBarras"
    
End Function

Public Sub Show()
    'Parent.Show
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

'***ALTERACAO POR TULIO EM 28/05***
Private Sub BotaoCancel_Click()

    'fecha a tela
    Unload Me

End Sub
'***FIM ALTERACAO POR TULIO EM 28/05***

'***ALTERACAO POR TULIO EM 28/05***
Private Sub BotaoRetornar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoRetornar_Click

    'Move o conteudo do grid para a memoria, ou seja, carrega a colecao de codigo de barras com os dados do grid
    lErro = Move_GridCodBarras_Memoria()
    If lErro <> SUCESSO Then gError 101706

    'fecha a tela
    Unload Me

    Exit Sub
    
Erro_BotaoRetornar_Click:

    Select Case gErr
    
        Case 101706
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165572)
        
    End Select

    Exit Sub

End Sub
'***FIM ALTERACAO POR TULIO EM 28/05***

'***ALTERACAO POR TULIO EM 28/05***
'CRIADA A FUNCAO
Private Function Move_GridCodBarras_Memoria() As Long
'Move o conteudo do grid para a colecao de codigo de barras...
'colCodBarras eh parametro de OUTPUT que retorna o conteudo do grid (conjunto de codigos de barra)

Dim iLinha As Integer
Dim sCodigo As String
Dim iIndice As Integer

On Error GoTo Erro_Move_GridCodBarras_Memoria

    'limpa o conteudo da colecao
    Set gobjProduto.colCodBarras = New Collection

    'para cada linha do grid
    For iLinha = 1 To objGridCodBarras.iLinhasExistentes
    
        'se o codigo de barras da linha corrente estiver preenchido
        If Len(Trim(GridCodBarras.TextMatrix(iLinha, iGrid_CodBarras_Col))) > 0 Then
            If Len(Trim(GridCodBarras.TextMatrix(iLinha, iGrid_CodBarras_Col))) > PRODUTO_CODBARRAS_MAX Or Len(Trim(GridCodBarras.TextMatrix(iLinha, iGrid_CodBarras_Col))) < PRODUTO_CODBARRAS_MIN Then gError 112415
            
            For iIndice = 1 To gobjProduto.colCodBarras.Count
                sCodigo = gobjProduto.colCodBarras.Item(iIndice)
                If sCodigo = GridCodBarras.TextMatrix(iLinha, iGrid_CodBarras_Col) Then gError 112541
            Next
            
            'adiciona na colecao o codigo da linha corrente
            gobjProduto.colCodBarras.Add GridCodBarras.TextMatrix(iLinha, iGrid_CodBarras_Col)
            
        End If
    
    Next
        
    Move_GridCodBarras_Memoria = SUCESSO
    
    Exit Function

Erro_Move_GridCodBarras_Memoria:

    Move_GridCodBarras_Memoria = gErr

    Select Case gErr
        
        Case 112415
            Call Rotina_Erro(vbOKOnly, "ERRO_CODBARRA_DIFERE_NUM_CARACTERES", gErr, PRODUTO_CODBARRAS_MAX, PRODUTO_CODBARRAS_MIN)
            
        Case 112541
            Call Rotina_Erro(vbOKOnly, "ERRO_CODBARRA_REPITIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165573)
            
    End Select

    Exit Function

End Function
'***FIM ALTERACAO POR TULIO EM 28/05***

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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'***ALTERACAO POR TULIO EM 28/05***
Private Function Inicializa_GridCodBarras(objGridInt As AdmGrid) As Long
'Inicializa o grid da tela

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridCodBarras

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Código de barras")
    
    'campos de edição do grid
    objGridInt.colCampo.Add (CodBarras.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_CodBarras_Col = 1
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridCodBarras

    'Numero Maximo de Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_CODBARRAS_PRODUTO

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = LINHAS_VISIVEIS_GRIDCODBARRAS

    'Largura da primeira coluna
    GridCodBarras.ColWidth(0) = 0

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_GridCodBarras = SUCESSO

    Exit Function

Erro_Inicializa_GridCodBarras:

    Inicializa_GridCodBarras = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165574)
            
    End Select

    Exit Function

End Function

'******************************************
'4 eventos do controle do GridCodBarras: CodBarras
'******************************************

Private Sub CodBarras_Change()

End Sub

Private Sub CodBarras_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridCodBarras)

End Sub

Private Sub CodBarras_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCodBarras)

End Sub
Private Sub CodBarras_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCodBarras.objControle = CodBarras
    lErro = Grid_Campo_Libera_Foco(objGridCodBarras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'******************************************
'Fim 4 eventos do controle do GridCodBarras: CodBarras
'******************************************

'*****************************************
'9 eventos do gridCodBarras que devem ser tratados
'*****************************************

Public Sub GridCodBarras_Click()

Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Click(objGridCodBarras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCodBarras, iAlterado)
    End If

End Sub

Public Sub GridCodBarras_GotFocus()
    Call Grid_Recebe_Foco(objGridCodBarras)
End Sub

Public Sub GridCodBarras_EnterCell()
    
Dim iAlterado As Integer
    
    Call Grid_Entrada_Celula(objGridCodBarras, iAlterado)
        
End Sub

Public Sub GridCodBarras_LeaveCell()
    
    Call Saida_Celula(objGridCodBarras)

End Sub

Public Sub GridCodBarras_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCodBarras)
    
End Sub

Public Sub GridCodBarras_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
Dim iAlterado As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCodBarras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then Call Grid_Entrada_Celula(objGridCodBarras, iAlterado)
    
End Sub

Public Sub GridCodBarras_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridCodBarras)

End Sub

Public Sub GridCodBarras_RowColChange()
    
    Call Grid_RowColChange(objGridCodBarras)
    
End Sub

Public Sub GridCodBarras_Scroll()
    
    Call Grid_Scroll(objGridCodBarras)

End Sub

'************************************************
'fim dos 9 eventos do gridCodBarras que devem ser tratados
'************************************************

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do gridCodBarras
'que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'inicializa a saida
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
        
        'Verifica qual celula do grid esta deixando
        'de ser a corrente para chamar a funcao de
        'saida celula adequada...
        Select Case objGridInt.objGrid.Col

            'se for a celula de Regiao de Venda
            Case iGrid_CodBarras_Col
        
                'chama a funcao adequada para o tratamento...
                lErro = Saida_Celula_CodBarras(objGridInt)
                If lErro <> SUCESSO Then gError 101709
            
        End Select
            
    End If

    'finaliza a saida
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101710
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 101709
        
        Case 101710
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165575)

    End Select

    Exit Function

End Function

Private Function ValidaDVEAN(ByVal sEAN As String) As Long

Dim intPar As Integer
Dim intImpar As Integer
Dim intSomaPar As Integer
Dim intSomaImpar As Integer
Dim intTotalSoma As Integer
Dim intDv As Integer
Dim i As Integer, iTam As Integer, iFatorPar As Integer, iFatorImpar As Integer, bPar As Boolean

On Error GoTo Erro_ValidaDVEAN

    sEAN = Trim(sEAN)
    iTam = Len(sEAN)

    If (iTam <> 14 And iTam <> 13 And iTam <> 12 And iTam <> 8) Or Not IsNumeric(sEAN) Then gError 201585

    intSomaPar = 0
    intSomaImpar = 0
    intTotalSoma = 0
    intDv = 0
    
    For i = 1 To (iTam - 1)
    
        bPar = (i Mod 2 = 0)
    
        If bPar Then
            intPar = CInt(Mid(sEAN, i, 1))
            intSomaPar = intSomaPar + intPar
        Else
            intImpar = CInt(Mid(sEAN, i, 1))
            intSomaImpar = intSomaImpar + intImpar
        End If
                
    Next

    Select Case iTam
    
        Case 8, 12, 14
        
            iFatorPar = 1
            iFatorImpar = 3
            
        Case 13
        
            iFatorPar = 3
            iFatorImpar = 1
    
    End Select
    
    intSomaPar = intSomaPar * iFatorPar
    intSomaImpar = intSomaImpar * iFatorImpar
    
    intTotalSoma = intSomaPar + intSomaImpar
    
    Do While intTotalSoma Mod 10 <> 0
    
        intDv = intDv + 1
        intTotalSoma = intTotalSoma + 1
    
    Loop
    
    If right(sEAN, 1) <> CStr(intDv) Then gError 201584
    
    ValidaDVEAN = SUCESSO
    
    Exit Function
    
Erro_ValidaDVEAN:

    ValidaDVEAN = gErr

    Select Case gErr
    
        Case 201584, 201585

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201583)

    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_CodBarras(objGridInt As AdmGrid) As Long
'Faz a crítica do campo CodBarras
'que está deixando de ser o campo corrente

Dim lErro As Long, sEAN As String, iTam As Integer

On Error GoTo Erro_Saida_Celula_CodBarras

    'instancia objcontrole como o controle Cliente
    Set objGridInt.objControle = CodBarras

    iTam = Len(Trim(objGridInt.objControle.Text))
    
    If iTam <> 0 Then

        
        'Verifica se o codigo informado ja aparece em outra linha do grid
        lErro = VerificaRepeteco_CodBarras(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 101713
    
        'se nao tiver de 8 a 14 caracteres, erro
        If iTam > PRODUTO_CODBARRAS_MAX Or iTam < PRODUTO_CODBARRAS_MIN Then gError 101739
        
        sEAN = Trim(objGridInt.objControle.Text)
            
        'validar ean
        lErro = ValidaDVEAN(sEAN)
        If lErro <> SUCESSO Then gError 201586
    
    
    End If
    
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101712

    'adiciona a linha se a corrente for a ultima e seu campo estiver preenchido
    If iTam <> 0 Then Call Adiciona_Linha_Seguinte

    Saida_Celula_CodBarras = SUCESSO

    Exit Function

Erro_Saida_Celula_CodBarras:

    Saida_Celula_CodBarras = gErr

    Select Case gErr


        Case 201586
            Call Rotina_Erro(vbOKOnly, "ERRO_EAN13_INVALIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 101712, 101713, 201586
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 101739
            Call Rotina_Erro(vbOKOnly, "ERRO_CODBARRA_EXCEDE_NUM_CARACTERES", gErr, objGridInt.objControle.Text, Len(Trim(objGridInt.objControle.Text)), PRODUTO_CODBARRAS_MIN, PRODUTO_CODBARRAS_MAX)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165576)

    End Select

    Exit Function

End Function

Private Function VerificaRepeteco_CodBarras(sCodBarras As String) As Long
'Verifica se existe o codigo que esta em sCodBarras
'mais de uma vez no grid
'sCodBarras eh parametro de INPUT e traz o codigo a ser verificado

Dim iLinha As Integer
Dim iPrimeiraLinhaRepetida As Integer

'parametro que denota a quantidade de codigos repetidos encontradas
'no grid
Dim iQtdCod As Integer

On Error GoTo Erro_VerificaRepeteco_CodBarras

    'para cada linha do grid
    For iLinha = 1 To objGridCodBarras.iLinhasExistentes
    
        'se o codigo de barras da linha corrente for = ao codigo passado como parametro
        If GridCodBarras.TextMatrix(iLinha, iGrid_CodBarras_Col) = sCodBarras Then
    
            'incrementa a qtd de codigos
            iQtdCod = iQtdCod + 1
            
            'se a qtd de codigos repetida for = 2, significa q achei a 1a linha
            'repetida
            'armazeno isso para poder enviar uma msg de erro + elaborada..
            If iQtdCod = 2 Then
                iPrimeiraLinhaRepetida = iLinha
            End If
            
    
        End If
    
    Next
    
    'se a qtd de codigos for maior do q 1, significa que existem linhas repetidas no grid
    If iQtdCod > 1 Then gError 101714
    
    VerificaRepeteco_CodBarras = SUCESSO

    Exit Function

Erro_VerificaRepeteco_CodBarras:

    VerificaRepeteco_CodBarras = gErr
    
    Select Case gErr
    
        Case 101714
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_REPETIDA_GRIDCODBARRAS", gErr, iQtdCod, iPrimeiraLinhaRepetida)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165577)
            
    End Select
    
    Exit Function

End Function

Private Sub Adiciona_Linha_Seguinte()
'adiciona a linha se a corrente for a
'ultima e se a coluna corrente estiver preenchida

    'se for ultima linha do grid habilitada e o campo estiver preenchido
    If GridCodBarras.Row - GridCodBarras.FixedRows = objGridCodBarras.iLinhasExistentes And Len(Trim(GridCodBarras.TextMatrix(GridCodBarras.Row, GridCodBarras.Col))) > 0 Then
        
        'inclui a proxima linha
        objGridCodBarras.iLinhasExistentes = objGridCodBarras.iLinhasExistentes + 1

    End If

End Sub

'***FIM ALTERACAO POR TULIO EM 28/05***
