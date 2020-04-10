VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl Grade 
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   ScaleHeight     =   5670
   ScaleWidth      =   7815
   Begin VB.ComboBox Layout 
      Height          =   315
      ItemData        =   "Grade.ctx":0000
      Left            =   1230
      List            =   "Grade.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2895
      Width           =   3375
   End
   Begin VB.TextBox Categoria 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   3720
      Width           =   1665
   End
   Begin VB.ComboBox Posicao 
      Height          =   315
      ItemData        =   "Grade.ctx":001D
      Left            =   3045
      List            =   "Grade.ctx":001F
      TabIndex        =   15
      Top             =   3720
      Width           =   1050
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5325
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Grade.ctx":0021
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "Grade.ctx":017B
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Grade.ctx":0305
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Grade.ctx":0837
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Grades 
      Height          =   4545
      Left            =   4905
      TabIndex        =   5
      Top             =   960
      Width           =   2700
   End
   Begin VB.ListBox Categorias 
      Height          =   1410
      Left            =   1245
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1410
      Width           =   3360
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1245
      TabIndex        =   0
      Top             =   495
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1245
      TabIndex        =   1
      Top             =   960
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridCategoria 
      Height          =   2160
      Left            =   1215
      TabIndex        =   4
      Top             =   3390
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   3810
      _Version        =   393216
      Rows            =   3
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
   End
   Begin VB.Label Label10 
      Caption         =   "Layout:"
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
      Height          =   315
      Left            =   465
      TabIndex        =   17
      Top             =   2940
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Grades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   4890
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Categorias:"
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
      Index           =   1
      Left            =   195
      TabIndex        =   8
      Top             =   1395
      Width           =   975
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   975
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   510
      TabIndex        =   6
      Top             =   525
      Width           =   660
   End
End
Attribute VB_Name = "Grade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Grid Categoria
Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Posicao_Col As Integer

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'*************************************************************************
'******************** INICIALIZAÇÃO DA TELA ******************************
'*************************************************************************
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa a variavel do grid
    Set objGridCategoria = New AdmGrid

    'Inicializa o grid
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError 123224

    'Carrega a ListBox de Categorias com as Categorias existentes no BD
    lErro = Carrega_Lista_Categorias()
    If lErro <> SUCESSO Then gError 86268

    'Carrega a ListBox de Grades com as Grades existentes no BD
    lErro = Carrega_Lista_Grades()
    If lErro <> SUCESSO Then gError 86269

    Call Carrega_Combo_Posicao

    Layout.ListIndex = 0

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 86268, 86269, 123224

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161654)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Categoria(objGridCategoria As AdmGrid) As Long
'realiza a inicialização do grid

    'tela em questão
    Set objGridCategoria.objForm = Me

    'titulos do grid
    objGridCategoria.colColuna.Add (" ")
    objGridCategoria.colColuna.Add ("Categoria")
    objGridCategoria.colColuna.Add ("Posição")

   'campos de edição do grid
    objGridCategoria.colCampo.Add (Categoria.Name)
    objGridCategoria.colCampo.Add (Posicao.Name)

    'atribui valor as colunas
    iGrid_Categoria_Col = 1
    iGrid_Posicao_Col = 2

    objGridCategoria.objGrid = GridCategoria

    'todas as linhas do grid
    objGridCategoria.objGrid.Rows = 101

    objGridCategoria.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'linhas visiveis do grid
    objGridCategoria.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridCategoria.ColWidth(0) = 0

    'Largura automática para as outras colunas
    objGridCategoria.iGridLargAuto = GRID_LARGURA_MANUAL

    'Proibido incluir no grid
    'objGridCategoria.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Proibido excluir no grid
    objGridCategoria.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGridCategoria)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Function Saida_Celula(objGridCategoria As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridCategoria)

    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridCategoria.objGrid.Col

            'Coluna Posição
            Case iGrid_Posicao_Col
                lErro = Saida_Celula_Posicao(objGridCategoria)
                If lErro <> SUCESSO Then gError 123221

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridCategoria)
        If lErro <> SUCESSO Then gError 123222

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 123221, 123222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161655)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Posicao(objGridCategoria As AdmGrid) As Long
'Faz a critica da célula Posicao do grid

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_Posicao

    Set objGridCategoria.objControle = Posicao

    'Verifica se alterou a linha atual
    If GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Posicao_Col) <> Posicao.Text Then

        iLinha = 1

        'se a linha foi a primeira
        'Modifica a posição da segunda
        If GridCategoria.Row = 1 Then iLinha = 2
        
        'Se a posição for Linha inclui no grid
        If Posicao.Text = Posicao.List(0) Then

            GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Posicao_Col) = Posicao.List(1)

        Else
        'Se a posição for coluna inclui no grid
        
            GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Posicao_Col) = Posicao.List(0)

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridCategoria)
    If lErro <> SUCESSO Then Error 123223

    Saida_Celula_Posicao = SUCESSO

    Exit Function

Erro_Saida_Celula_Posicao:

    Saida_Celula_Posicao = gErr

    Select Case gErr

        Case 123223
            Call Grid_Trata_Erro_Saida_Celula(objGridCategoria)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161656)

    End Select

    Exit Function

End Function

Function Carrega_Lista_Grades() As Long
'Preenche a ListBox de Grades

Dim lErro As Long
Dim objGrade As ClassGrade
Dim colGrade As New Collection

On Error GoTo Erro_Carrega_Lista_Grades

    'Lê todas as Grades de Produto
    lErro = CF("Grades_Le_Todas", colGrade)
    If lErro <> SUCESSO Then gError 86259

    'Adiciona as Grades lidas na List
    For Each objGrade In colGrade
        Grades.AddItem objGrade.sCodigo
    Next

    Carrega_Lista_Grades = SUCESSO

    Exit Function

Erro_Carrega_Lista_Grades:

    Carrega_Lista_Grades = gErr

    Select Case gErr

        Case 86259

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161657)

    End Select

    Exit Function

End Function

Function Carrega_Lista_Categorias() As Long
'Preenche a ListBox de Categorias

Dim lErro As Long
Dim objCategoriaProduto As ClassCategoriaProduto
Dim colCategoriaProduto As New Collection

On Error GoTo Erro_Carrega_Lista_Categorias

    'Lê todas as Categorias de Produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO Then gError 86258

    'Adiciona as categorias lidas na List
    For Each objCategoriaProduto In colCategoriaProduto
        Categorias.AddItem objCategoriaProduto.sCategoria
    Next

    Carrega_Lista_Categorias = SUCESSO

    Exit Function

Erro_Carrega_Lista_Categorias:

    Carrega_Lista_Categorias = gErr

    Select Case gErr

        Case 86258

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161658)

    End Select

    Exit Function

End Function

Private Sub Carrega_Combo_Posicao()
'Preenche a ComboPosição com Linha e Coluna

    Posicao.AddItem ("Linha")
    Posicao.ItemData(Posicao.NewIndex) = 0
    Posicao.AddItem ("Coluna")
    Posicao.ItemData(Posicao.NewIndex) = 1

End Sub

'*************************************************************************
'******************** PASSAGEM DE PARÂMETRO PARA A TELA ******************
'*************************************************************************
Function Trata_Parametros(Optional objGrade As ClassGrade) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se algum parâmetro foi passado
    If Not (objGrade Is Nothing) Then

        'Verifica se o Código veio preenchido
        If Len(Trim(objGrade.sCodigo)) > 0 Then

            'Tenta ler Grade com o código passado
            lErro = CF("Grade_Le", objGrade)
            If lErro <> SUCESSO And lErro <> 86275 Then gError 86270

            'Se encontrou a grade
            If lErro = SUCESSO Then

                'Traz a Grade para a tela
                lErro = Traz_Grade_Tela(objGrade)
                If lErro <> SUCESSO Then gError 86271

            'Senão
            Else

                'Coloca o Código passado na Tela
                Codigo.Text = objGrade.sCodigo

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 86270, 86271

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161659)

    End Select

    Exit Function

End Function

Function Traz_Grade_Tela(objGrade As ClassGrade) As Long
'Coloca na Telas Os dados da Grade passada por parâmetro

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objGradeCategoria As ClassGradeCategoria

On Error GoTo Erro_Traz_Grade_Tela

    Call Limpa_Tela_Grade

    'Lê as Categorias relacionadas a essa grade
    lErro = CF("GradeCategoria_Le", objGrade)
    If lErro <> SUCESSO Then gError 86276

    'Coloca o Código da Grade na Tela
    Codigo.Text = objGrade.sCodigo

    'coloca a Descrição da Grade na tela
    Descricao.Text = objGrade.sDescricao
    
    Call Combo_Seleciona_ItemData(Layout, objGrade.iLayout)
    
    'Para cada Categoria lida
    For Each objGradeCategoria In objGrade.colCategoria

        'Percorre a lista de Categorias em busca da categoria lida
        For iIndice2 = 0 To Categorias.ListCount - 1

            'Se encontrar a categoria
            If objGradeCategoria.sCategoria = Categorias.List(iIndice2) Then

                'Seleciona na lista a categoria
                Categorias.Selected(iIndice2) = True
                
                '???? Escreve a posicao dessa categoria
                
                'Escreve a posicao dessa categoria no grid
                For iIndice = 0 To 1
                
                    If objGradeCategoria.iPosicao = Posicao.ItemData(iIndice) Then
                    
                        GridCategoria.TextMatrix(objGridCategoria.iLinhasExistentes, iGrid_Posicao_Col) = Posicao.List(iIndice)
                    
                        Exit For
                        
                    End If
                    
                Next
                
                'Termina a busca por essa categoria
                Exit For

            End If

        Next

    Next

    iAlterado = 0

    Traz_Grade_Tela = SUCESSO

    Exit Function

Erro_Traz_Grade_Tela:

    Traz_Grade_Tela = gErr

    Select Case gErr

        Case 86276

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161660)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EMPENHO
    Set Form_Load_Ocx = Me
    Caption = "Grade"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Grade"

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

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 122694

    'Limpa a tela
    Call Limpa_Tela_Grade

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 122694

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161661)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Grade()

Dim iIndice As Integer

    'Limpa o codigo e a descrição
    Call Limpa_Tela(Me)

    'Limpa o grid categoria
    Call Grid_Limpa(objGridCategoria)
    
    Layout.ListIndex = 0

    'Limpa a ListBox de categoria
    For iIndice = 0 To Categorias.ListCount - 1
        Categorias.Selected(iIndice) = False
    Next

    iAlterado = 0

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Verifica se dados necessários da grade foram preenchidos
'Atualiza/Insere Grade no BD
'Atualiza List

Dim lErro As Long
Dim iIndice As Integer
Dim objGrade As New ClassGrade

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 122695

    'Verifica se a Descricao foi preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 122696

    'Move os dados da tela para ObjGrade
    lErro = Move_Tela_Memoria(objGrade)
    If lErro <> SUCESSO Then gError 122697

    'Caso o numero de categorias selecionadas seja diferente de 1 e 2 -> ERRO
    'If objGrade.colCategoria.Count <> 1 And objGrade.colCategoria.Count <> 2 Then gError 122698

    'DEPOIS IRÁ SUBIR
    lErro = CF("Grade_Grava", objGrade)
    If lErro <> SUCESSO Then gError 122699

    'Retira a Grade contida em objGrade da Lista de Grades (caso seja uma atualização)
    Call Retira_Grade_Lista(objGrade)

    'Adiciona a Grade contida em objGrade à lista de Grades
    Grades.AddItem (objGrade.sCodigo)

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 122695
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 122696
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 122697, 122699

        Case 122698
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_CATEGORIAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161662)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name

        Case Posicao.Name

            'Se existirem duas linhas no grid preenchida libera a ComboPosição
            If objGridCategoria.iLinhasExistentes >= 2 Then
                Posicao.Enabled = True
            Else
                Posicao.Enabled = False
            End If

        End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161663)

    End Select

    Exit Sub

End Sub

Sub Preenche_Grid(iItem As Integer)
'Preenche o GridCategoria com os elementos selecionados na ListCategoria

Dim iIndice As Integer
Dim iQuantidade As Integer
Dim asItens(0 To 10) As String

    'Realiza uma contagem de quantos elementos foram selecionados na ListCategoria
    For iIndice = 0 To Categorias.ListCount - 1

        'Se o elemento está selecionado incrementa o número de quantidade e armzena-o em um array
        If Categorias.Selected(iIndice) = True Then

            asItens(iQuantidade) = Categorias.List(iIndice)
            iQuantidade = iQuantidade + 1

        End If

    Next

    'Se o Item recebido como referência foi selecionado
    If Categorias.Selected(iItem) = True Then

        'Se o grid já possui duas linhas existentes --> Sai da rotina
        'If objGridCategoria.iLinhasExistentes = 2 Then Exit Sub

        'Coloca o elemento selecionado no grid
        objGridCategoria.iLinhasExistentes = objGridCategoria.iLinhasExistentes + 1

        GridCategoria.TextMatrix(objGridCategoria.iLinhasExistentes, iGrid_Categoria_Col) = Categorias.Text

        If objGridCategoria.iLinhasExistentes = 1 Then

            GridCategoria.TextMatrix(objGridCategoria.iLinhasExistentes, iGrid_Posicao_Col) = Posicao.List(1)

        Else

            GridCategoria.TextMatrix(objGridCategoria.iLinhasExistentes, iGrid_Posicao_Col) = Posicao.List(0)

        End If

    Else

        'Verifica os elementos existentes no grid excluindo ou incluindo
        For iIndice = 1 To objGridCategoria.iLinhasExistentes

            'Se o elemento da linha atual for iqual ao elemento que não está selecionado então exclui do grid
            If GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = Categorias.List(iItem) Then

                Call Grid_Exclui_Linha(objGridCategoria, iIndice)

            End If

        Next
        'Verifica a quantidade de linhas existentes
        'Se a quantidade de elemento selecionados na List for > 1 então inclui um novo elemento
        If objGridCategoria.iLinhasExistentes = 1 Then

            GridCategoria.TextMatrix(1, iGrid_Posicao_Col) = Posicao.List(1)

            If iQuantidade > 1 Then

                For iIndice = 0 To iQuantidade

                    'Inclui um novo elemento selecionado
                    If GridCategoria.TextMatrix(objGridCategoria.iLinhasExistentes, iGrid_Categoria_Col) <> asItens(iIndice) Then

                        objGridCategoria.iLinhasExistentes = objGridCategoria.iLinhasExistentes + 1

                        GridCategoria.TextMatrix(objGridCategoria.iLinhasExistentes, iGrid_Categoria_Col) = asItens(iIndice)
                        GridCategoria.TextMatrix(objGridCategoria.iLinhasExistentes, iGrid_Posicao_Col) = Posicao.List(1)

                        Exit For

                    End If

                Next

            End If

        End If

    End If

   Exit Sub

End Sub

Sub Retira_Grade_Lista(objGrade As ClassGrade)
'Retira a grade contida em Objgrade da lista de Grades, caso ela pertença à lista

Dim iIndice As Integer

    'Percorre a lista de Grades em busca da Grade lida
    For iIndice = 0 To Grades.ListCount - 1

        'Se encontrar a grade com mesmo código da que está armazenada no obj
        If objGrade.sCodigo = Grades.List(iIndice) Then

            'Remove a grade da lista de grades
            Grades.RemoveItem (iIndice)

        End If
    Next

End Sub

Function Move_Tela_Memoria(objGrade As ClassGrade) As Long
'Carrega em objGrade os dados da tela

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim iIndice3 As Integer
Dim objGradeCategoria As ClassGradeCategoria
Dim bTemColuna As Boolean

On Error GoTo Erro_Move_Tela_Memoria

    bTemColuna = False

    objGrade.sCodigo = Codigo.Text

    objGrade.sDescricao = Descricao.Text
    
    objGrade.iLayout = Layout.ItemData(Layout.ListIndex)

    'Move cada categoria selecionada e a posição objGrade
    For iIndice = 0 To Categorias.ListCount - 1

        'Se a categoria estiver selecionada então
        If Categorias.Selected(iIndice) = True Then

            Set objGradeCategoria = New ClassGradeCategoria
            
            'Move a categoria para o objGradeCategoria
            objGradeCategoria.sCategoria = Categorias.List(iIndice)
            
            '????? Genérico
            For iIndice3 = 1 To objGridCategoria.iLinhasExistentes
    
                'Se a Categoria do grid for igual então preenche a posição
                If GridCategoria.TextMatrix(iIndice3, iGrid_Categoria_Col) = Categorias.List(iIndice) Then
                
                    'Preenche a posição da Categoria Selecionada no objGradeCategoria
                    For iIndice2 = 0 To 1
                        
                        If GridCategoria.TextMatrix(iIndice3, iGrid_Posicao_Col) = Posicao.List(iIndice2) Then
                        
                            objGradeCategoria.iPosicao = Posicao.ItemData(iIndice2)
                            objGradeCategoria.iSeq = iIndice3
                        End If
                    
                    Next
            
                End If
                
                If objGradeCategoria.iPosicao = 1 Then bTemColuna = True
                
            Next
                
            'Adiciona na coleção o objGradeCategoria
            objGrade.colCategoria.Add objGradeCategoria
                
        End If

    Next
    
    If Not bTemColuna Then gError 180105

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 180105
            Call Rotina_Erro(vbOKOnly, "ERRO_GRADE_SEM_COLUNA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161664)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objGrade As New ClassGrade
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 122727

    'Carrega em objGrade os dados da tela
    lErro = Move_Tela_Memoria(objGrade)
    If lErro <> SUCESSO Then gError 122729

    'Verifica se a Grade de objGrade existe no BD
    lErro = CF("Grade_Le", objGrade)
    If lErro <> SUCESSO And lErro <> 86275 Then gError 122747

    'Se não achou a grade no BD -> Erro
    If lErro = 86275 Then gError 122728

    'Pede confirmação para exclusão da Grade
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_GRADE", objGrade.sCodigo)

    'Sai da rotina caso não seja seja confirmada a exclusão
    If vbMsgRes = vbNo Then Exit Sub

    'Exclui a Grade passada em objGrade
    lErro = CF("Grade_Exclui", objGrade)
    If lErro <> SUCESSO Then gError 122730

    'Retira a Grade com o Código armazenado em objGrade da lista de grades
    Call Retira_Grade_Lista(objGrade)

    'Limpa a tela
    Call Limpa_Tela_Grade

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 122727
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 122728
            Call Rotina_Erro(vbOKOnly, "ERRO_GRADE_NAO_CADASTRADA", gErr, objGrade.sCodigo)

        Case 122729, 122730, 122747

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161665)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 86280

    Call Limpa_Tela_Grade

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 86281

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 86280, 86281

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161666)

    End Select

    Exit Sub

End Sub

Private Sub Grades_Click()

Dim lErro As Long
Dim objGrade As New ClassGrade

On Error GoTo Erro_Grades_Click

    '???? Se
    '????ListIndex = -1
    
    'Se o índice do elemento selecionado for -1 sai da rotina
    If Grades.ListIndex = -1 Then Exit Sub

    objGrade.sCodigo = Grades.List(Grades.ListIndex)

    lErro = CF("Grade_Le", objGrade)
    If lErro <> SUCESSO And lErro <> 86275 Then gError 86277
    If lErro <> SUCESSO Then gError 86278

    lErro = Traz_Grade_Tela(objGrade)
    If lErro <> SUCESSO Then gError 86279

    iAlterado = 0

    Exit Sub

Erro_Grades_Click:

    Select Case gErr

        Case 86277, 86279

        Case 86278
            Call Rotina_Erro(vbOKOnly, "ERRO_GRADE_NAO_CADASTRADA", gErr, objGrade.sCodigo)

            Grades.RemoveItem Grades.ListIndex

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161667)

    End Select

    Exit Sub

End Sub

'*************************************************************************
'****************** TRATAMENTO DE CONTROLES DA TELA **********************
'*************************************************************************

Private Sub Categorias_ItemCheck(Item As Integer)

    iAlterado = REGISTRO_ALTERADO

    Call Preenche_Grid(Item)

End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

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

        iAlterado = REGISTRO_ALTERADO

    Exit Sub

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

Private Sub GridCategoria_RowColChange()

    Call Grid_RowColChange(objGridCategoria)

End Sub

Private Sub GridCategoria_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Private Sub Categoria_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Categoria_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Private Sub Categoria_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Private Sub Categoria_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = Categoria
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Posicao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Posicao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Private Sub Posicao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Private Sub Posicao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = Posicao
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridCategoria = Nothing

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objGrade As New ClassGrade

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Grade"

    'Le os dados da Tela Grade
    lErro = Move_Tela_Memoria(objGrade)
    If lErro <> SUCESSO Then gError 122745

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objGrade.sCodigo, STRING_GRADE_CODIGO, "Codigo"
    colCampoValor.Add "Descricao", objGrade.sDescricao, STRING_GRADE_DESCRICAO, "Descricao"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 122745

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161668)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objGrade As New ClassGrade

On Error GoTo Erro_Tela_Preenche

    objGrade.sCodigo = colCampoValor.Item("Codigo").vValor
    objGrade.sDescricao = colCampoValor.Item("Descricao").vValor

    'Traz dados da Grade para a Tela
    lErro = Traz_Grade_Tela(objGrade)
    If lErro <> SUCESSO Then gError 122746

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 122746

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161669)

    End Select

    Exit Sub

End Sub
