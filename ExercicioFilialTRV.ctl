VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ExercicioFilial 
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   ScaleHeight     =   4200
   ScaleWidth      =   7590
   Begin VB.ComboBox FechadoCTB 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "ExercicioFilialTRV.ctx":0000
      Left            =   5460
      List            =   "ExercicioFilialTRV.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1710
      Width           =   1080
   End
   Begin VB.ComboBox Fechado 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "ExercicioFilialTRV.ctx":001F
      Left            =   4365
      List            =   "ExercicioFilialTRV.ctx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1725
      Width           =   1080
   End
   Begin MSMask.MaskEdBox NomePeriodo 
      Height          =   285
      Left            =   1755
      TabIndex        =   16
      Top             =   1740
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
      Left            =   3105
      TabIndex        =   17
      Tag             =   "1"
      Top             =   1740
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
   Begin VB.ComboBox Exercicio 
      Height          =   315
      ItemData        =   "ExercicioFilialTRV.ctx":003E
      Left            =   1050
      List            =   "ExercicioFilialTRV.ctx":0040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   270
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   1665
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "ExercicioFilialTRV.ctx":0042
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "ExercicioFilialTRV.ctx":019C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "ExercicioFilialTRV.ctx":06CE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridPeriodos 
      Height          =   1860
      Left            =   675
      TabIndex        =   5
      Top             =   1635
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
   Begin VB.Label DataFimExercicio 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6135
      TabIndex        =   14
      Top             =   885
      Width           =   1290
   End
   Begin VB.Label DataInicioExercicio 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3585
      TabIndex        =   13
      Top             =   885
      Width           =   1290
   End
   Begin VB.Label NomeExterno 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   750
      TabIndex        =   12
      Top             =   885
      Width           =   1425
   End
   Begin VB.Label Label3 
      Caption         =   "Exercicio:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   315
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
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
      Left            =   2430
      TabIndex        =   10
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3135
      TabIndex        =   9
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label LabelDataInicio 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicio:"
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
      Left            =   2505
      TabIndex        =   8
      Top             =   930
      Width           =   1005
   End
   Begin VB.Label LabelDataFim 
      AutoSize        =   -1  'True
      Caption         =   "Data Fim:"
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
      Left            =   5265
      TabIndex        =   7
      Top             =   930
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   930
      Width           =   555
   End
End
Attribute VB_Name = "ExercicioFilial"
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
Const GRID_STATUSCTB_COL = 4
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
    If lErro <> SUCESSO Then gError 207112

    lErro = Inicializa_Grid_ExercicioTela(objGrid1)
    If lErro <> SUCESSO Then gError 207113
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 207112, 207113

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207114)

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
    objGridInt.colColuna.Add ("StatusCTB")

   'campos de edição do grid
    objGridInt.colCampo.Add (NomePeriodo.Name)
    objGridInt.colCampo.Add (DataInicioPeriodo.Name)
    objGridInt.colCampo.Add (Fechado.Name)
    objGridInt.colCampo.Add (FechadoCTB.Name)

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
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then gError 207115

    For Each objExercicio In colExercicios
    
        'preenche ComboBox com NomeExterno e ItemData com Exercicio
        Exercicio.AddItem CStr(objExercicio.iExercicio)

    Next

    Preenche_ComboExercicio = SUCESSO

    Exit Function

Erro_Preenche_ComboExercicio:

    Preenche_ComboExercicio = gErr

    Select Case gErr

        Case 207115

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207116)

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
        If iIndice = Exercicio.ListCount Then gError 207117
    
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 207117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_ENCONTRADO_TELA", gErr, iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207118)

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

        If GridPeriodos.TextMatrix(iIndice, GRID_STATUSCTB_COL) = EXERCICIO_DESC_FECHADO Then
            objPeriodo.iFechadoCTB = PERIODO_FECHADO
        Else
            objPeriodo.iFechadoCTB = PERIODO_ABERTO
        End If


        colPeriodos.Add objPeriodo

    Next
    
    MoveDadosTela_Variaveis = SUCESSO

    Exit Function

Erro_MoveDadosTela_Variaveis:

    MoveDadosTela_Variaveis = gErr

    Select Case gErr

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207119)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'Limpa todos os campos da Tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    If Status.Caption <> "" Then

        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then gError 207120

        Call Limpa_Tela_ExercicioTela
        
        iAlterado = 0
        
    End If

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 207120

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207121)

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
    If lErro <> SUCESSO Then gError 207122

    'pega o valor do novo exercicio
    iExercicio2 = CInt(Exercicio.Text)

    'exibe na tela os dados para o exercício atual
    lErro = Carregar_Dados_Tela(iExercicio2)
    If lErro <> SUCESSO Then gError 207123

    GridPeriodos.TopRow = 1
    
    iAlterado = 0

    Exit Sub

Erro_Exercicio_Click:

    Select Case gErr

        Case 207122
            For iIndice = 0 To Exercicio.ListCount - 1
                If Exercicio.ItemData(iIndice) = iExercicio2 Then
                    Exercicio.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 207123

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207124)

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
            gError 207125

    End Select

    Status_Exibe_Descricao = SUCESSO

    Exit Function

Erro_Status_Exibe_Descricao:

    Status_Exibe_Descricao = gErr

    Select Case gErr

        Case 207125
             lErro = Rotina_Erro(vbOKOnly, "ERRO_STATUS_EXERCICIO_INVALIDO", gErr, iStatus)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207126)

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
    lErro = CF("Exercicio_Le", iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then gError 207127

    'se o exercicio não está cadastrado ==> erro
    If lErro = 10083 Then gError 207128

    objExerciciosFilial.iFilialEmpresa = giFilialEmpresa
    objExerciciosFilial.iExercicio = iExercicio

    'Le o exerciciofilial passado como parametro
    lErro = CF("ExerciciosFilial_Le", objExerciciosFilial)
    If lErro <> SUCESSO And lErro <> 20389 Then gError 207129

    'se o exerciciofilial não está cadastrado ==> erro
    If lErro = 20389 Then gError 207130

    'exibe descrição do status
    lErro = Status_Exibe_Descricao(objExerciciosFilial.iStatus, sDescricaoStatus)
    If lErro <> SUCESSO Then gError 207131

    Status.Caption = sDescricaoStatus
    
    NomeExterno.Caption = objExercicio.sNomeExterno

    'exibe data início e data fim
    DataInicioExercicio.Caption = Format(objExercicio.dtDataInicio, "dd/mm/yyyy")
    DataFimExercicio.Caption = Format(objExercicio.dtDataFim, "dd/mm/yyyy")

    'lê todos os períodos para o exercício determinado
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then gError 207132

    'preenche o grid com os períodos lidos
    lErro = PreencheGridPeriodos(colPeriodos)
    If lErro <> SUCESSO Then gError 207133

    iAlterado = 0

    Carregar_Dados_Tela = SUCESSO

    Exit Function

Erro_Carregar_Dados_Tela:

    Carregar_Dados_Tela = gErr

    Select Case gErr

        Case 207127, 207129, 207131, 207132, 207133

        Case 207128
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", gErr, iExercicio)

        Case 207130
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOFILIAL_NAO_CADASTRADO", gErr, iExercicio, giFilialEmpresa)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207134)

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

        If objPeriodo.iFechadoCTB = PERIODO_FECHADO Then
            GridPeriodos.TextMatrix(iIndice, GRID_STATUSCTB_COL) = EXERCICIO_DESC_FECHADO
        Else
            GridPeriodos.TextMatrix(iIndice, GRID_STATUSCTB_COL) = EXERCICIO_DESC_ABERTO
        End If


    Next

    PreencheGridPeriodos = SUCESSO

    Exit Function

Erro_PreencheGridPeriodos:

    PreencheGridPeriodos = gErr

    Select Case gErr

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207136)

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
    If lErro <> SUCESSO Then gError 207138

    'Limpa a tela
    Call Limpa_Tela_ExercicioTela

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 207138

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207139)

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
    If Exercicio.ListIndex = -1 Then gError 207140

    'move os dados da tela para as variáveis
    lErro = MoveDadosTela_Variaveis(objExercicio, colPeriodos)
    If lErro <> SUCESSO Then gError 207141

    lErro = CF("ExerciciosFilial_Grava", objExercicio, colPeriodos)
    If lErro <> SUCESSO Then gError 207142
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 207140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_SELECIONADO", gErr)

        Case 207141, 207142

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207143)

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
                lErro = Saida_Celula_Fechado(objGridInt)
                If lErro <> SUCESSO Then gError 207144

            Case GRID_STATUSCTB_COL
                lErro = Saida_Celula_FechadoCTB(objGridInt)
                If lErro <> SUCESSO Then gError 207145

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 207146

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 207144, 107145

        Case 207146
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207147)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Fechado(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Fechado

    Set objGridInt.objControle = Fechado

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 207148

    Saida_Celula_Fechado = SUCESSO

    Exit Function

Erro_Saida_Celula_Fechado:

    Saida_Celula_Fechado = gErr

    Select Case gErr

        Case 207148
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207149)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FechadoCTB(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FechadoCTB

    Set objGridInt.objControle = FechadoCTB

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 207150

    Saida_Celula_FechadoCTB = SUCESSO

    Exit Function

Erro_Saida_Celula_FechadoCTB:

    Saida_Celula_FechadoCTB = gErr

    Select Case gErr

        Case 207150
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207152)

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

Private Sub FechadoCTB_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FechadoCTB_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FechadoCTB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub FechadoCTB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub FechadoCTB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = FechadoCTB
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub
