VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ExercicioFilialOcx 
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   LockControls    =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   7605
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1665
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "ExercicioFilialOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "ExercicioFilialOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "ExercicioFilialOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.ComboBox Fechado 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "ExercicioFilialOcx.ctx":080A
      Left            =   4785
      List            =   "ExercicioFilialOcx.ctx":0814
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1590
      Width           =   1080
   End
   Begin VB.ComboBox Exercicio 
      Height          =   315
      ItemData        =   "ExercicioFilialOcx.ctx":0829
      Left            =   1050
      List            =   "ExercicioFilialOcx.ctx":082B
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   780
   End
   Begin MSMask.MaskEdBox NomePeriodo 
      Height          =   285
      Left            =   2175
      TabIndex        =   1
      Top             =   1605
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   20
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
   Begin MSMask.MaskEdBox DataInicioPeriodo 
      Height          =   285
      Left            =   3525
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1605
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridPeriodos 
      Height          =   1860
      Left            =   1110
      TabIndex        =   4
      Top             =   1590
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   120
      Top             =   915
      Width           =   555
      ForeColor       =   128
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
   End
   Begin VB.Label LabelDataFim 
      Height          =   195
      Left            =   5265
      Top             =   915
      Width           =   825
      ForeColor       =   128
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      Caption         =   "Data Fim:"
   End
   Begin VB.Label LabelDataInicio 
      Height          =   195
      Left            =   2505
      Top             =   915
      Width           =   1005
      ForeColor       =   128
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      Caption         =   "Data Inicio:"
   End
   Begin VB.Label Status 
      Height          =   315
      Left            =   3135
      Top             =   255
      Width           =   1095
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label4 
      Height          =   195
      Left            =   2430
      Top             =   285
      Width           =   615
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      Caption         =   "Status:"
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      Top             =   300
      Width           =   855
      ForeColor       =   128
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exercicio:"
   End
   Begin VB.Label NomeExterno 
      Height          =   315
      Left            =   750
      Top             =   870
      Width           =   1425
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label DataInicioExercicio 
      Height          =   315
      Left            =   3585
      Top             =   870
      Width           =   1290
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label DataFimExercicio 
      Height          =   315
      Left            =   6135
      Top             =   870
      Width           =   1290
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "ExercicioFilialOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Const GRID_NOME_COL = 1
Const GRID_DATAINI_COL = 2
Const GRID_STATUS_COL = 3
Dim iIndiceAtual As Integer
Dim iExercicioMudou As Integer
Dim objGrid1 As AdmGrid
Dim iExercicio2 As Integer

Private Sub Exercicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fechado_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0
    
    iExercicio2 = 0

    Set objGrid1 = New AdmGrid

    'inicializa o combo com os exercícios existentes no BD
    lErro = Preenche_ComboExercicio()
    If lErro <> SUCESSO Then Error 10226

    lErro = Inicializa_Grid_ExercicioTela(objGrid1)
    If lErro <> SUCESSO Then Error 10227
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 10226, 10227

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159757)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Inicializa_Grid_ExercicioTela(objGridInt As AdmGrid) As Long
'inicializa o grid de períodos do form ExercicioTela

Dim lErro As Long

   Set objGrid1.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("Periodo")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Data Inicio")
    objGridInt.colColuna.Add ("Status")

   'campos de edição do grid
    objGridInt.colCampo.Add (NomePeriodo.Name)
    objGridInt.colCampo.Add (DataInicioPeriodo.Name)
    objGridInt.colCampo.Add (Fechado.Name)

    objGridInt.objGrid = GridPeriodos

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 13

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    GridPeriodos.ColWidth(0) = 1000

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR

    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_ExercicioTela = SUCESSO

End Function

Function Preenche_ComboExercicio() As Long
'preenche Combo de Exercicios

Dim colExercicios As New Collection
Dim lErro As Long
Dim iConta As Integer
Dim objExercicio As ClassExercicio

On Error GoTo Erro_Preenche_ComboExercicio

    Exercicio.Clear

    'le todos os exercícios existentes no BD
    lErro = CF("Exercicios_Le_Todos",colExercicios)
    If lErro <> SUCESSO Then Error 10228

    For Each objExercicio In colExercicios
    
        'preenche ComboBox com NomeExterno e ItemData com Exercicio
        Exercicio.AddItem CStr(objExercicio.iExercicio)

    Next

    Preenche_ComboExercicio = SUCESSO

    Exit Function

Erro_Preenche_ComboExercicio:

    Preenche_ComboExercicio = Err

    Select Case Err

        Case 10228

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159758)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objExercicio As ClassExercicio) As Long

Dim lErro As Long
Dim iExercicio As Integer
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'verifica se o exercício não está preenchido
    If Not (objExercicio Is Nothing) Then
        iExercicio = objExercicio.iExercicio
    
        For iIndice = 0 To Exercicio.ListCount - 1
            If Exercicio.List(iIndice) = CStr(iExercicio) Then
                Exercicio.ListIndex = iIndice
                Exit For
            End If
        Next
        
        'se não encontrou o exercicio na combobox
        If iIndice = Exercicio.ListCount Then Error 10229
    
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 10229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_ENCONTRADO_TELA", Err, iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159759)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function MoveDadosTela_Variaveis(objExercicio As ClassExercicio, colPeriodos As Collection) As Long
'Move os dados do exercicio da tela para objExercicio e os
'dados referebtes ao período para colPeriodos

Dim iIndice As Integer
Dim objPeriodo As ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_MoveDadosTela_Variaveis

    'dados do exercício
    objExercicio.iExercicio = iExercicio2
    
    objExercicio.iFilialEmpresa = giFilialEmpresa

    'dados dos períodos
    For iIndice = 1 To objGrid1.iLinhasExistentes

        Set objPeriodo = New ClassPeriodo
        
        If GridPeriodos.TextMatrix(iIndice, GRID_STATUS_COL) = EXERCICIO_DESC_FECHADO Then
            objPeriodo.iFechado = PERIODO_FECHADO
        Else
            objPeriodo.iFechado = PERIODO_ABERTO
        End If

        colPeriodos.Add objPeriodo

    Next
    
    MoveDadosTela_Variaveis = SUCESSO

    Exit Function

Erro_MoveDadosTela_Variaveis:

    MoveDadosTela_Variaveis = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159760)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'Limpa todos os campos da Tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    If Status.Caption <> "" Then

        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then Error 10230

        Call Limpa_Tela_ExercicioTela
        
        iAlterado = 0
        
    End If

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 10230

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159761)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_ExercicioTela()

    'Limpa os Campos da Tela
    DataInicioExercicio.Caption = ""
    DataFimExercicio.Caption = ""
    Status.Caption = ""
    Exercicio.ListIndex = -1
    NomeExterno.Caption = ""

    Call Grid_Limpa(objGrid1)
    
    GridPeriodos.TopRow = 1
    
    iExercicio2 = 0

End Sub

Private Sub Exercicio_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Exercicio_Click

    If Exercicio.ListIndex = -1 Then Exit Sub

    If Exercicio.Text = CStr(iExercicio2) Then Exit Sub
    
    'verifica se existe a necessidade de salvar o exercício antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 10231

    'pega o valor do novo exercicio
    iExercicio2 = CInt(Exercicio.Text)

    'exibe na tela os dados para o exercício atual
    lErro = Carregar_Dados_Tela(iExercicio2)
    If lErro <> SUCESSO Then Error 10232

    GridPeriodos.TopRow = 1
    
    iAlterado = 0

    Exit Sub

Erro_Exercicio_Click:

    Select Case Err

        Case 10231
            For iIndice = 0 To Exercicio.ListCount - 1
                If Exercicio.ItemData(iIndice) = iExercicio2 Then
                    Exercicio.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 10232

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159762)

    End Select

    Exit Sub

End Sub

Private Sub Fechado_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
     
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGrid1 = Nothing
    
End Sub

Private Sub NomePeriodo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid1)
End Sub

Private Sub DataInicioPeriodo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
End Sub

Private Sub DataInicioPeriodo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid1)
End Sub

Function Status_Exibe_Descricao(iStatus As Integer, sDescricao As String) As Long
'retorna em sDescricao a descrição do status iStatus
'1-Aberto, 2-Apurado, 3-Fechado

Dim lErro As Long

On Error GoTo Erro_Status_Exibe_Descricao

    Select Case iStatus

        Case EXERCICIO_ABERTO
            sDescricao = EXERCICIO_DESC_ABERTO
        Case EXERCICIO_APURADO
            sDescricao = EXERCICIO_DESC_APURADO
        Case EXERCICIO_FECHADO
            sDescricao = EXERCICIO_DESC_FECHADO
        Case Else
            Error 10233

    End Select

    Status_Exibe_Descricao = SUCESSO

    Exit Function

Erro_Status_Exibe_Descricao:

    Status_Exibe_Descricao = Err

    Select Case Err

        Case 10233
             lErro = Rotina_Erro(vbOKOnly, "ERRO_STATUS_EXERCICIO_INVALIDO", Err, iStatus)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159763)

    End Select

    Exit Function

End Function

Function Carregar_Dados_Tela(iExercicio As Integer) As Long
'carrega os dados para a tela referente ao exercício 'iExercicio'

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim sDescricaoStatus As String
Dim colPeriodos As New Collection
Dim objExerciciosFilial As New ClassExerciciosFilial

On Error GoTo Erro_Carregar_Dados_Tela

    'Le o exercicio passado como parametro
    lErro = CF("Exercicio_Le",iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 10234

    'se o exercicio não está cadastrado ==> erro
    If lErro = 10083 Then Error 10235

    objExerciciosFilial.iFilialEmpresa = giFilialEmpresa
    objExerciciosFilial.iExercicio = iExercicio

    'Le o exerciciofilial passado como parametro
    lErro = CF("ExerciciosFilial_Le",objExerciciosFilial)
    If lErro <> SUCESSO And lErro <> 20389 Then Error 55826

    'se o exerciciofilial não está cadastrado ==> erro
    If lErro = 20389 Then Error 55827

    'exibe descrição do status
    lErro = Status_Exibe_Descricao(objExerciciosFilial.iStatus, sDescricaoStatus)
    If lErro <> SUCESSO Then Error 10236

    Status.Caption = sDescricaoStatus
    
    NomeExterno.Caption = objExercicio.sNomeExterno

    'exibe data início e data fim
    DataInicioExercicio.Caption = Format(objExercicio.dtDataInicio, "dd/mm/yyyy")
    DataFimExercicio.Caption = Format(objExercicio.dtDataFim, "dd/mm/yyyy")

    'lê todos os períodos para o exercício determinado
    lErro = CF("Periodo_Le_Todos_Exercicio",giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 10237

    'preenche o grid com os períodos lidos
    lErro = PreencheGridPeriodos(colPeriodos)
    If lErro <> SUCESSO Then Error 10238

    iAlterado = 0

    Carregar_Dados_Tela = SUCESSO

    Exit Function

Erro_Carregar_Dados_Tela:

    Carregar_Dados_Tela = Err

    Select Case Err

        Case 10234, 10236, 10237, 10238, 55826

        Case 10235
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, iExercicio)

        Case 55827
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOFILIAL_NAO_CADASTRADO", Err, iExercicio, giFilialEmpresa)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159764)

    End Select

    Exit Function

End Function

Function PreencheGridPeriodos(colPeriodos As Collection) As Long
'preenche o grid com os períodos passados na coleção colPeriodos

Dim lErro As Long
Dim iIndice As Integer
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheGridPeriodos

    'Limpa o grid
    Call Grid_Limpa(objGrid1)

    objGrid1.iLinhasExistentes = colPeriodos.Count

    'preenche o grid com os dados retornados na coleção colPeriodos
    For iIndice = 1 To colPeriodos.Count

        Set objPeriodo = colPeriodos.Item(iIndice)
        
        GridPeriodos.TextMatrix(iIndice, GRID_NOME_COL) = objPeriodo.sNomeExterno
        GridPeriodos.TextMatrix(iIndice, GRID_DATAINI_COL) = Format(objPeriodo.dtDataInicio, "dd/mm/yyyy")

        If objPeriodo.iFechado = PERIODO_FECHADO Then
            GridPeriodos.TextMatrix(iIndice, GRID_STATUS_COL) = EXERCICIO_DESC_FECHADO
        Else
            GridPeriodos.TextMatrix(iIndice, GRID_STATUS_COL) = EXERCICIO_DESC_ABERTO
        End If

    Next

    PreencheGridPeriodos = SUCESSO

    Exit Function

Erro_PreencheGridPeriodos:

    PreencheGridPeriodos = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159765)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub DataInicioPeriodo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = DataInicioPeriodo
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Fechado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Fechado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Fechado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Fechado
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NomePeriodo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
End Sub

Private Sub NomePeriodo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = NomePeriodo
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridPeriodos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridPeriodos_GotFocus()
    Call Grid_Recebe_Foco(objGrid1)
End Sub

Private Sub GridPeriodos_EnterCell()
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
End Sub

Private Sub GridPeriodos_LeaveCell()
    Call Saida_Celula(objGrid1)
End Sub

Private Sub GridPeriodos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
End Sub

Private Sub GridPeriodos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridPeriodos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGrid1)
End Sub

Private Sub GridPeriodos_RowColChange()
    Call Grid_RowColChange(objGrid1)
End Sub

Private Sub GridPeriodos_Scroll()
    Call Grid_Scroll(objGrid1)
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'grava o exercicio em questão
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 10252

    'Limpa a tela
    Call Limpa_Tela_ExercicioTela

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 10252

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159766)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'grava os dados do exercicio em questão

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim colPeriodos As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se existe pelo menos um exercicio preenchido
    If Exercicio.ListIndex = -1 Then Error 10253

    'move os dados da tela para as variáveis
    lErro = MoveDadosTela_Variaveis(objExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 10254

    lErro = CF("ExerciciosFilial_Grava",objExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 10255
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 10253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_SELECIONADO", Err)

        Case 10254, 10255

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159767)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case GRID_STATUS_COL
                lErro = Saida_Celula_Status(objGridInt)
                If lErro <> SUCESSO Then Error 10256

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 10257

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 10256

        Case 10257
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159768)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Status(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Status

    Set objGridInt.objControle = Fechado

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 10263

    Saida_Celula_Status = SUCESSO

    Exit Function

Erro_Saida_Celula_Status:

    Saida_Celula_Status = Err

    Select Case Err

        Case 10263
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159769)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EXERCICIO_FILIAL
    Set Form_Load_Ocx = Me
    Caption = "Exercício - Periodos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExercicioFilial"
    
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




Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelDataFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataFim, Source, X, Y)
End Sub

Private Sub LabelDataFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataFim, Button, Shift, X, Y)
End Sub

Private Sub LabelDataInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataInicio, Source, X, Y)
End Sub

Private Sub LabelDataInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataInicio, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub NomeExterno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NomeExterno, Source, X, Y)
End Sub

Private Sub NomeExterno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NomeExterno, Button, Shift, X, Y)
End Sub

Private Sub DataInicioExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataInicioExercicio, Source, X, Y)
End Sub

Private Sub DataInicioExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataInicioExercicio, Button, Shift, X, Y)
End Sub

Private Sub DataFimExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataFimExercicio, Source, X, Y)
End Sub

Private Sub DataFimExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataFimExercicio, Button, Shift, X, Y)
End Sub

