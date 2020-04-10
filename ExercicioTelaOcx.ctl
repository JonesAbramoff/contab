VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.UserControl ExercicioTelaOcx 
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   LockControls    =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   7815
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ExercicioTelaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ExercicioTelaOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ExercicioTelaOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ExercicioTelaOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Exercicio 
      Height          =   315
      ItemData        =   "ExercicioTelaOcx.ctx":0994
      Left            =   1095
      List            =   "ExercicioTelaOcx.ctx":0996
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "Geração  Automática  de  Periodos"
      Height          =   870
      Left            =   120
      TabIndex        =   18
      Top             =   1380
      Width           =   7545
      Begin VB.CommandButton BotaoGeraPeriodos 
         DisabledPicture =   "ExercicioTelaOcx.ctx":0998
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         Picture         =   "ExercicioTelaOcx.ctx":25DA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gerar Períodos"
         Top             =   255
         Width           =   1305
      End
      Begin VB.ComboBox Periodicidade 
         Height          =   315
         ItemData        =   "ExercicioTelaOcx.ctx":421C
         Left            =   1440
         List            =   "ExercicioTelaOcx.ctx":4235
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   345
         Width           =   1455
      End
      Begin MSMask.MaskEdBox NumPeriodos 
         Height          =   315
         Left            =   4710
         TabIndex        =   5
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown SpinNumPeriodos 
         Height          =   330
         Left            =   5055
         TabIndex        =   17
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelNumPeriodos 
         AutoSize        =   -1  'True
         Caption         =   "Num. de Periodos:"
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
         Left            =   3060
         TabIndex        =   19
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label LabelPeriodicidade 
         AutoSize        =   -1  'True
         Caption         =   "Periodicidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   390
         Width           =   1230
      End
   End
   Begin VB.TextBox NomeExterno 
      Height          =   315
      Left            =   810
      MaxLength       =   20
      TabIndex        =   1
      Top             =   870
      Width           =   1305
   End
   Begin MSMask.MaskEdBox NomePeriodo 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   2535
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptInclude   =   0   'False
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
   Begin MSMask.MaskEdBox DataInicioExercicio 
      Height          =   315
      Left            =   3690
      TabIndex        =   2
      Top             =   870
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSMask.MaskEdBox DataFimExercicio 
      Height          =   315
      Left            =   6195
      TabIndex        =   3
      Top             =   870
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSMask.MaskEdBox DataInicioPeriodo 
      Height          =   285
      Left            =   4110
      TabIndex        =   8
      Tag             =   "1"
      Top             =   2550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
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
      Left            =   1725
      TabIndex        =   9
      Top             =   2475
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   3
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSComCtl2.UpDown SpinDataInicio 
      Height          =   330
      Left            =   4890
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   855
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown SpinDataFim 
      Height          =   330
      Left            =   7395
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   870
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
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
      Left            =   150
      TabIndex        =   21
      Top             =   300
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
      Left            =   2460
      TabIndex        =   22
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3150
      TabIndex        =   23
      Top             =   255
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
      Left            =   2535
      TabIndex        =   24
      Top             =   915
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
      Left            =   5295
      TabIndex        =   25
      Top             =   915
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
      Left            =   150
      TabIndex        =   26
      Top             =   915
      Width           =   555
   End
End
Attribute VB_Name = "ExercicioTelaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Codigos de Periodicidade de Exercicio
Const PERIODICIDADE_ANUAL = 1
Const PERIODICIDADE_BIMENSAL = 2
Const PERIODICIDADE_LIVRE = 3
Const PERIODICIDADE_MENSAL = 4
Const PERIODICIDADE_QUADRIMESTRAL = 5
Const PERIODICIDADE_SEMESTRAL = 6
Const PERIODICIDADE_TRIMESTRAL = 7

Dim iAlterado As Integer
Const GRID_NOME_COL = 1
Const GRID_DATAINI_COL = 2
Dim objGrid1 As AdmGrid
Dim iExercicio2 As Integer

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Exercicio.ListIndex = -1 Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_EXERCICIO")
    
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Chama Rotina que exclui o exercicio
    lErro = CF("Exercicio_Exclui", CInt(Exercicio.Text))
    If lErro <> SUCESSO Then Error 11415

    'inicializa o combo com os exercícios existentes no BD
    lErro = Preenche_ComboExercicio()
    If lErro <> SUCESSO Then Error 10130
    
    'Limpa a tela
    Call Limpa_Tela_ExercicioTela
        
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 10130, 11415

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159770)

    End Select

    Exit Sub

End Sub

Private Sub DataFimExercicio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFimExercicio, iAlterado)

End Sub

Private Sub DataInicioExercicio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicioExercicio, iAlterado)

End Sub

Private Sub Exercicio_Change()

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
    If lErro <> SUCESSO Then Error 13637

    lErro = Inicializa_Grid_ExercicioTela(objGrid1)
    If lErro <> SUCESSO Then Error 13636

    'Colocar periodicidade inicial = LIVRE
    For iIndice = 0 To Periodicidade.ListCount - 1
        If Periodicidade.ItemData(iIndice) = PERIODICIDADE_MENSAL Then
            Periodicidade.ListIndex = iIndice
            Exit For
        End If
    Next
    Call HabilitaCampos

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 13636, 13637

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159771)

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

   'campos de edição do grid
    objGridInt.colCampo.Add (NomePeriodo.Name)
    objGridInt.colCampo.Add (DataInicioPeriodo.Name)

    objGridInt.objGrid = GridPeriodos

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 13

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    GridPeriodos.ColWidth(0) = 1000

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

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
    If lErro <> SUCESSO Then Error 13635

    For Each objExercicio In colExercicios
    
        'preenche ComboBox com NomeInterno
        Exercicio.AddItem CStr(objExercicio.iExercicio)

    Next

    'insere um novo exercicio
    Exercicio.AddItem CStr(Exercicio.ListCount + 1)

    Preenche_ComboExercicio = SUCESSO

    Exit Function

Erro_Preenche_ComboExercicio:

    Preenche_ComboExercicio = Err

    Select Case Err

        Case 13635

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159772)

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
        If iIndice = Exercicio.ListCount Then Error 10131
    
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 10131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_ENCONTRADO_TELA", Err, iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159773)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub GeracaoPeriodos(dtDataInicio As Date, dtDataFim As Date, iPeriodicidade As Integer, iNumPeriodo As Integer, colPeriodos As Collection)
'gera períodos entre as datas de entrada ( dtDataInicio e dtDataFim )
'os períodos são retornados na coleção colPeriodos

Dim iDuracaoPeriodo As Integer

    'determina a duração de cada período
    Select Case iPeriodicidade

        Case PERIODICIDADE_ANUAL
            iDuracaoPeriodo = 12
        Case PERIODICIDADE_BIMENSAL
            iDuracaoPeriodo = 2
        Case PERIODICIDADE_MENSAL
            iDuracaoPeriodo = 1
        Case PERIODICIDADE_QUADRIMESTRAL
            iDuracaoPeriodo = 4
        Case PERIODICIDADE_SEMESTRAL
            iDuracaoPeriodo = 6
        Case PERIODICIDADE_TRIMESTRAL
            iDuracaoPeriodo = 3
        Case PERIODICIDADE_LIVRE
            'Calcula a duracao de periodos de acordo com a quantidade de periodos passada como parametro
            Call CalculaPeriodos_Livre(dtDataInicio, dtDataFim, iNumPeriodo, colPeriodos)
            Exit Sub
    End Select

    'Traz em ColPeriodos todos os periodos calculados
    Call Calcula_Periodos(dtDataInicio, dtDataFim, iDuracaoPeriodo, colPeriodos)

    Exit Sub

End Sub

Private Sub HabilitaCampos()
'só pode gerar períodos se a periodicidade não for livre
'para periodicidade livre os períodos são determinados por Número de Períodos

Dim iHabBotaoGera As Integer
Dim iHabNumPer As Integer
Dim iHabDataExercicioInicio As Integer
Dim iHabDataExercicioFim As Integer

    iHabBotaoGera = False
    iHabDataExercicioFim = False
    iHabDataExercicioInicio = False
    iHabNumPer = False

    'Exercício selecionado e Data Início e Data Fim preenchidos
    If objGrid1.iProibidoIncluir = 0 And Exercicio.ListIndex <> -1 And Len(DataInicioExercicio.ClipText) > 0 And Len(DataFimExercicio.ClipText) > 0 Then
    
        iHabBotaoGera = True

        'periodicidade livre permite selecionar o numero de períodos a serem gerados
        If Periodicidade.ItemData(Periodicidade.ListIndex) = PERIODICIDADE_LIVRE Then
            iHabNumPer = True
        End If
        
    End If

    'se tiver um exercicio selecionado
    If Exercicio.ListIndex <> -1 Then

        'se só tiver um exercicio ou tiverem 2 e estiver editando o primeiro
        If (Exercicio.ListCount = 2 And Exercicio.ListIndex = 0) Or Exercicio.ListCount = 1 Then
            iHabDataExercicioFim = True
            iHabDataExercicioInicio = True
            
        ElseIf Exercicio.ListIndex = Exercicio.ListCount - 2 Or Exercicio.ListIndex = Exercicio.ListCount - 1 Then
            iHabDataExercicioFim = True

        End If
        
    End If
    
    'habilita ou desabilita os campos
    BotaoGeraPeriodos.Enabled = iHabBotaoGera
    DataInicioPeriodo.Enabled = iHabBotaoGera
    NumPeriodos.Enabled = iHabNumPer
    SpinNumPeriodos.Enabled = iHabNumPer
    DataInicioExercicio.Enabled = iHabDataExercicioInicio
    SpinDataInicio.Enabled = iHabDataExercicioInicio
    DataFimExercicio.Enabled = iHabDataExercicioFim
    SpinDataFim.Enabled = iHabDataExercicioFim

End Sub

Function MoveDadosTela_Variaveis(objExercicio As ClassExercicio, colPeriodos As Collection) As Long
'Move os dados do exercicio da tela para objExercicio e os
'dados referebtes ao período para colPeriodos

Dim iIndice As Integer
Dim objPeriodo As ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_MoveDadosTela_Variaveis

    'dados do exercício
    objExercicio.sNomeExterno = NomeExterno.Text
    objExercicio.dtDataInicio = CDate(DataInicioExercicio.Text)
    objExercicio.dtDataFim = CDate(DataFimExercicio.Text)
    objExercicio.iNumPeriodos = objGrid1.iLinhasExistentes
    objExercicio.iExercicio = iExercicio2

    'dados dos períodos
    For iIndice = 1 To objGrid1.iLinhasExistentes

        Set objPeriodo = New ClassPeriodo
        
        objPeriodo.sNomeExterno = GridPeriodos.TextMatrix(iIndice, GRID_NOME_COL)
        If objPeriodo.sNomeExterno = "" Then Error 11584

        If GridPeriodos.TextMatrix(iIndice, GRID_DATAINI_COL) = "" Then Error 10100
        objPeriodo.dtDataInicio = CDate(GridPeriodos.TextMatrix(iIndice, GRID_DATAINI_COL))

        'se for o primeiro periodo, verifica se a data inicio coincide com a data inicio do exercicio.
        If iIndice = 1 And objPeriodo.dtDataInicio <> objExercicio.dtDataInicio Then Error 10134

        If iIndice = objGrid1.iLinhasExistentes Then
            objPeriodo.dtDataFim = DataFimExercicio.Text
            If objExercicio.dtDataFim < objPeriodo.dtDataInicio Then Error 55725
        Else
            If GridPeriodos.TextMatrix(iIndice + 1, GRID_DATAINI_COL) = "" Then Error 10101
            objPeriodo.dtDataFim = (CDate(GridPeriodos.TextMatrix(iIndice + 1, GRID_DATAINI_COL)) - 1)
        End If

        objPeriodo.iFechado = PERIODO_ABERTO
        objPeriodo.iFechadoCTB = PERIODO_ABERTO

        colPeriodos.Add objPeriodo

    Next

    MoveDadosTela_Variaveis = SUCESSO

    Exit Function

Erro_MoveDadosTela_Variaveis:

    MoveDadosTela_Variaveis = Err

    Select Case Err

        Case 10100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA", Err, iIndice)

        Case 10101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA", Err, iIndice + 1)

        Case 10134
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINI_PRIMEIRO_PERIODO", Err, CStr(objPeriodo.dtDataInicio), CStr(objExercicio.dtDataInicio))

        Case 11584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_PERIODO_VAZIO", Err)

        Case 55725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_ULT_PERIODO_FORA_EXERCICIO", Err, objPeriodo.dtDataInicio, objExercicio.dtDataInicio, objExercicio.dtDataFim)
            
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159774)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'Limpa todos os campos da Tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    If Status.Caption <> "" Then

        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then Error 11587

        Call Limpa_Tela_ExercicioTela
        
        iAlterado = 0
        
    End If

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 11587

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159775)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_ExercicioTela()

    'Limpa os Campos da Tela
    DataInicioExercicio.PromptInclude = False
    DataInicioExercicio.Text = ""
    DataInicioExercicio.PromptInclude = True
    DataFimExercicio.PromptInclude = False
    DataFimExercicio.Text = ""
    DataFimExercicio.PromptInclude = True
    Periodicidade.ListIndex = -1
    NumPeriodos.PromptInclude = False
    NumPeriodos.Text = ""
    NumPeriodos.PromptInclude = True
    Status.Caption = ""
    Exercicio.ListIndex = -1
    NomeExterno.Text = ""

    Call Grid_Limpa(objGrid1)
    
    Call HabilitaCampos
    
    GridPeriodos.TopRow = 1
    
    iExercicio2 = 0

End Sub

Private Sub DataFimExercicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoGeraPeriodos_Click()

Dim lErro As Long
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo
Dim iConta As Integer
Dim dtDataFim As Date
Dim dtDataInicio As Date
Dim iPeriodicidade As Integer
Dim iNumPeriodo As Integer

On Error GoTo Erro_BotaoGeraPeriodos_Click

    dtDataInicio = CDate(DataInicioExercicio.Text)
    dtDataFim = CDate(DataFimExercicio.Text)
    
    iPeriodicidade = Periodicidade.ItemData(Periodicidade.ListIndex)

    If Trim(NumPeriodos.Text) = "" Then
        iNumPeriodo = 1
        NumPeriodos.PromptInclude = False
        NumPeriodos.Text = "1"
        NumPeriodos.PromptInclude = True
    Else
        iNumPeriodo = CInt(NumPeriodos.Text)
    End If

    'Chama a Rotina que gera todos os periodos
    Call GeracaoPeriodos(dtDataInicio, dtDataFim, iPeriodicidade, iNumPeriodo, colPeriodos)

    'preenche o grid com os períodos gerados
    lErro = PreencheGridPeriodos(colPeriodos)
    If lErro <> SUCESSO Then Error 13622

    GridPeriodos.TopRow = 1

    Exit Sub

Erro_BotaoGeraPeriodos_Click:

    Select Case Err

        Case 13622

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159776)

    End Select

    Exit Sub

End Sub

Private Sub DataFimExercicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFimExercicio_Validate

    'verifica se a data final está vazia
    If Len(DataFimExercicio.ClipText) = 0 Then Error 13623

    'verifica se a data final é válida
    lErro = Data_Critica(DataFimExercicio.Text)
    If lErro <> SUCESSO Then Error 11416

    'data inicial não pode ser maior que a data final
    If CDate(DataInicioExercicio.Text) > CDate(DataFimExercicio.Text) Then Error 13034

    Call HabilitaCampos

    Exit Sub

Erro_DataFimExercicio_Validate:

    Cancel = True


    Select Case Err

        Case 11416

        Case 13034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_EXERCICIO_MENOR", Err, DataFimExercicio.Text, DataInicioExercicio.Text)

        Case 13623
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_EXERCICIO_NAO_PREENCHIDA", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159777)

    End Select

    Exit Sub

End Sub

Private Sub DataInicioExercicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataInicioExercicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicioExercicio_Validate

    'verifica se a data inicial está vazia
    If Len(DataInicioExercicio.ClipText) = 0 Then Error 13624

    'verifica se a data inicial é válida
    lErro = Data_Critica(DataInicioExercicio.Text)
    If lErro <> SUCESSO Then Error 11417

    'data inicial não pode ser maior que a data final
    If CDate(DataInicioExercicio.Text) > CDate(DataFimExercicio.Text) Then Error 55726

    Call HabilitaCampos

    Exit Sub

Erro_DataInicioExercicio_Validate:

    Cancel = True


    Select Case Err

        Case 11417

        Case 13624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_EXERCICIO_NAO_PREENCHIDA", Err)

        Case 55726
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_EXERCICIO_MAIOR", Err, DataInicioExercicio.Text, DataFimExercicio.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159778)

    End Select

    Exit Sub

End Sub

Private Sub Exercicio_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Exercicio_Click

    If Exercicio.ListIndex = -1 Then Exit Sub

    If Exercicio.Text = CStr(iExercicio2) Then Exit Sub
    
    'verifica se existe a necessidade de salvar o exercício antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 13666

    'pega o valor do novo exercicio
    iExercicio2 = CInt(Exercicio.Text)

    'se selecionou o ultimo exercicio ==> está querendo criar um novo exercicio
    If Exercicio.ListIndex = Exercicio.ListCount - 1 Then
    
        lErro = Criar_Exercicio()
        If lErro <> SUCESSO Then Error 10075
        
    Else
    
        'exibe na tela os dados para o exercício atual
        lErro = Carregar_Dados_Tela(iExercicio2)
        If lErro <> SUCESSO Then Error 13625

        'verifica se existe movimento, lote ou lançamento pendente para o novo exercício
        lErro = CF("Exercicio_Critica_Movimento", iExercicio2)
        If lErro <> SUCESSO And lErro <> 13663 Then Error 13626

        'existe pendência, então desabilita campos ( data início e fim, incluindo periodicidade e o grid ) e sai
        If lErro = 13663 Then

            Call Trata_Exercicio_Com_Movimento
            
        Else
        
            Call Trata_Exercicio_Sem_Movimento
            
        End If
            
    End If

    GridPeriodos.TopRow = 1
    
    iAlterado = 0

    Exit Sub

Erro_Exercicio_Click:

    Select Case Err

        Case 10075, 13625, 13626

        Case 13666
            For iIndice = 0 To Exercicio.ListCount - 1
                If Exercicio.List(iIndice) = CStr(iExercicio2) Then
                    Exercicio.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159779)

    End Select

    Exit Sub

End Sub

Private Sub Trata_Exercicio_Com_Movimento()
'Se o exercicio possui movimento só pode alterar o nome do exercicio e dos periodos

    DataInicioExercicio.Enabled = False
    SpinDataInicio.Enabled = False
    DataFimExercicio.Enabled = False
    SpinDataFim.Enabled = False
    Periodicidade.Enabled = False
    NumPeriodos.Enabled = False
    SpinNumPeriodos.Enabled = False
    BotaoGeraPeriodos.Enabled = False
    DataInicioPeriodo.Enabled = False
    objGrid1.iProibidoIncluir = 1
    objGrid1.iProibidoExcluir = 1

End Sub

Private Sub Trata_Exercicio_Sem_Movimento()
'Se o exercicio não possui movimento só não pode alterar a data inicio do exercicio e a data fim (se não for o ultimo exercicio)

    objGrid1.iProibidoIncluir = 0
    objGrid1.iProibidoExcluir = 0
    BotaoGeraPeriodos.Enabled = True
    DataInicioPeriodo.Enabled = True
    Periodicidade.Enabled = True
    NumPeriodos.Enabled = True
    SpinNumPeriodos.Enabled = True
    DataInicioExercicio.Enabled = False
    SpinDataInicio.Enabled = False
    DataFimExercicio.Enabled = False
    SpinDataFim.Enabled = False
    
    'se for o ultimo exercicio cadastrado ==> pode alterar a data fim do exercicio
    If Exercicio.ListIndex = Exercicio.ListCount - 2 Then
        DataFimExercicio.Enabled = True
        SpinDataFim.Enabled = True
    End If
    
    If Exercicio.ListCount = 1 Or (Exercicio.ListIndex = 0 And Exercicio.ListCount = 2) Then
        DataInicioExercicio.Enabled = True
        SpinDataInicio.Enabled = True
        DataFimExercicio.Enabled = True
        SpinDataFim.Enabled = True
    End If
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGrid1 = Nothing
    
End Sub

Private Sub NomeExterno_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomePeriodo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomePeriodo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid1)
End Sub

Private Sub DataInicioPeriodo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
End Sub

Private Sub DataInicioPeriodo_Change()
    iAlterado = REGISTRO_ALTERADO
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
            Error 13627

    End Select

    Status_Exibe_Descricao = SUCESSO

    Exit Function

Erro_Status_Exibe_Descricao:

    Status_Exibe_Descricao = Err

    Select Case Err

        Case 13627
             lErro = Rotina_Erro(vbOKOnly, "ERRO_STATUS_EXERCICIO_INVALIDO", Err, iStatus)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159780)

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
    If lErro <> SUCESSO And lErro <> 10083 Then Error 13628

    'se o exercicio não está cadastrado ==> erro
    If lErro = 10083 Then Error 10088

    objExerciciosFilial.iFilialEmpresa = giFilialEmpresa
    objExerciciosFilial.iExercicio = iExercicio

    'Le o exerciciofilial passado como parametro
    lErro = CF("ExerciciosFilial_Le", objExerciciosFilial)
    If lErro <> SUCESSO And lErro <> 20389 Then Error 55828

    'se o exerciciofilial não está cadastrado ==> erro
    If lErro = 20389 Then Error 55829

    'exibe descrição do status
    lErro = Status_Exibe_Descricao(objExerciciosFilial.iStatus, sDescricaoStatus)
    If lErro <> SUCESSO Then Error 13629

    Status.Caption = sDescricaoStatus
    
    NomeExterno.Text = objExercicio.sNomeExterno

    'exibe data início e data fim
    DataInicioExercicio.Text = Format(objExercicio.dtDataInicio, "dd/mm/yy")
    DataFimExercicio.Text = Format(objExercicio.dtDataFim, "dd/mm/yy")

    'exibe número de períodos
    NumPeriodos.PromptInclude = False
    NumPeriodos.Text = CStr(objExercicio.iNumPeriodos)
    NumPeriodos.PromptInclude = True

    'exibe periodicidade livre
    lErro = SelecionaPeriodicidade(PERIODICIDADE_LIVRE)
    If lErro <> SUCESSO Then Error 13630
    Call HabilitaCampos

    'lê todos os períodos para o exercício determinado
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 13631

    'preenche o grid com os períodos lidos
    lErro = PreencheGridPeriodos(colPeriodos)
    If lErro <> SUCESSO Then Error 13632

    iAlterado = 0

    Carregar_Dados_Tela = SUCESSO

    Exit Function

Erro_Carregar_Dados_Tela:

    Carregar_Dados_Tela = Err

    Select Case Err

        Case 10088
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, iExercicio)
             
        Case 13628, 13629, 13630, 13631, 13632, 55828

        Case 55829
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOFILIAL_NAO_CADASTRADO", Err, iExercicio, giFilialEmpresa)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159781)

    End Select

    Exit Function

End Function

Function SelecionaPeriodicidade(iPeriodicidade As Integer) As Long
'seleciona periodicidade determinada por iPeriodicidade

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_SelecionaPeriodicidade

    'percorre todas as periodicidades
    For iIndice = 0 To Periodicidade.ListCount - 1

        'verifica se encontrou a periodicidade procurada
        If Periodicidade.ItemData(iIndice) = iPeriodicidade Then

            Periodicidade.ListIndex = iIndice
            Exit For

        End If
    Next

    'verifica se terminou a busca proque encontrou ou não
    If iIndice = Periodicidade.ListCount Then Error 13634

    SelecionaPeriodicidade = SUCESSO

    Exit Function

Erro_SelecionaPeriodicidade:

    SelecionaPeriodicidade = Err

    Select Case Err

        Case 13634
             lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODICIDADE_INVALIDA", Err, iPeriodicidade)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159782)

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

    Next

    PreencheGridPeriodos = SUCESSO

    Exit Function

Erro_PreencheGridPeriodos:

    PreencheGridPeriodos = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159783)

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
    If lErro <> SUCESSO Then Error 10073

    'inicializa o combo com os exercícios existentes no BD
    lErro = Preenche_ComboExercicio()
    If lErro <> SUCESSO Then Error 55727
    
    'Limpa a tela
    Call Limpa_Tela_ExercicioTela

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 10073, 55727

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159784)

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
    If Exercicio.ListIndex = -1 Then Error 11382

    'verifica se o nome do exercício foi preenchido
    If NomeExterno.Text = "" Then Error 10076

    'Verifica se pelo menos um periodo foi gerado
    If objGrid1.iLinhasExistentes = 0 Then Error 11381
    
    'move os dados da tela para as variáveis
    lErro = MoveDadosTela_Variaveis(objExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 11376

    lErro = CF("Exercicio_Grava", objExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 13667
    
    'Se acabou de inserir um novo exercicio, tem que inserir o proximo numero de exercicio.
    If Exercicio.List(Exercicio.ListCount - 1) = CStr(objExercicio.iExercicio) Then
        Exercicio.AddItem CStr(Exercicio.ListCount + 1)
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 10076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_EXERCICIO_VAZIO", Err)
            
        Case 11376, 13667

        Case 11381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_SEM_PERIODO", Err, NomeExterno.Text)

        Case 11382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_SELECIONADO", Err)
            
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159785)

    End Select

    Exit Function

End Function

Function Critica_Campo_DataInicioPeriodo(sData As String) As Long
'faz a crítica da data do início do período de acordo com o exercício e com o período anterior

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim dtData As Date

On Error GoTo Erro_Critica_Campo_DataInicioPeriodo

    lErro = Data_Critica(sData)
    If lErro <> SUCESSO Then Error 13638
    
    dtData = CDate(sData)

    objExercicio.dtDataInicio = CDate(DataInicioExercicio.Text)
    objExercicio.dtDataFim = CDate(DataFimExercicio.Text)

    'verifica se data dtData está dentro do exercício
    If dtData < objExercicio.dtDataInicio Or dtData > objExercicio.dtDataFim Then Error 13640

    'verifica se o período é maior que 1
    If GridPeriodos.Row > 1 Then

        'se a data da linha anterior estiver preenchida
        If Len(Trim(GridPeriodos.TextMatrix(GridPeriodos.Row - 1, GRID_DATAINI_COL))) > 0 Then

            'data início deve ser maior que a data início do período anterior
            If dtData <= CDate(GridPeriodos.TextMatrix(GridPeriodos.Row - 1, GRID_DATAINI_COL)) Then Error 13641
            
        End If

    Else

        'data início = data início exercício, quando período = 1
        If dtData <> objExercicio.dtDataInicio Then

            GridPeriodos.TextMatrix(GridPeriodos.Row, GRID_DATAINI_COL) = Format(objExercicio.dtDataInicio, "dd/mm/yyyy")
            Error 13642

        End If

    End If
    
    If GridPeriodos.Row < objGrid1.iLinhasExistentes Then
    
        'se a data do periodo seguinte estiver preenchido
        If Len(Trim(GridPeriodos.TextMatrix(GridPeriodos.Row + 1, GRID_DATAINI_COL))) > 0 Then
    
            If dtData >= CDate(GridPeriodos.TextMatrix(GridPeriodos.Row + 1, GRID_DATAINI_COL)) Then Error 10135
            
        End If

    End If

    Critica_Campo_DataInicioPeriodo = SUCESSO

    Exit Function

Erro_Critica_Campo_DataInicioPeriodo:

    Critica_Campo_DataInicioPeriodo = Err

    Select Case Err

        Case 10135
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINI_PERIODO_MAIOR_PERIODO_SEG", Err, Format(dtData, "dd/mm/yyyy"), GridPeriodos.TextMatrix(GridPeriodos.Row + 1, GRID_DATAINI_COL))

        Case 13638

        Case 13640
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FORA_EXERCICIO", Err, Format(dtData, "dd/mm/yyyy"), DataInicioExercicio.Text, DataFimExercicio.Text)

        Case 13641
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINI_PERIODO_MENOR_PERIODO_ANT", Err, Format(dtData, "dd/mm/yyyy"), GridPeriodos.TextMatrix(GridPeriodos.Row - 1, GRID_DATAINI_COL))

        Case 13642
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINI_PRIMEIRO_PERIODO", Err, Format(dtData, "dd/mm/yyyy"), DataInicioExercicio.Text)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159786)

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

            Case GRID_NOME_COL
                lErro = Saida_Celula_Nome(objGridInt)
                If lErro <> SUCESSO Then Error 11377

            Case GRID_DATAINI_COL
                lErro = Saida_Celula_DataIni(objGridInt)
                If lErro <> SUCESSO Then Error 11378

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 11380

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 11377, 11378

        Case 11380
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159787)

    End Select

    Exit Function

End Function

Private Sub NumPeriodos_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumPeriodos, iAlterado)

End Sub

Private Sub NumPeriodos_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iNumPeriodos As Integer

On Error GoTo Erro_NumPeriodos_Validate

    If Len(Trim(NumPeriodos.Text)) = 0 Then Exit Sub

    lErro = Valor_Positivo_Critica(NumPeriodos.Text)
    If lErro <> SUCESSO Then Error 57503
    
    iNumPeriodos = CInt(NumPeriodos.Text)

    'número de períodos não pode ser maior que NUM_MAX_PERIODOS
    If iNumPeriodos > NUM_MAX_PERIODOS Then

        NumPeriodos.Text = CStr(NUM_MAX_PERIODOS)
        Error 13679

    'número de períodos não pode ser menor que 1
    ElseIf iNumPeriodos < 1 Then

            NumPeriodos.Text = "1"
            Error 13680

    End If

    Exit Sub

Erro_NumPeriodos_Validate:

    Cancel = True


    Select Case Err

        Case 13679, 13680
             lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_PERIODO_INVALIDO", Err, NUM_MAX_PERIODOS)

        Case 57503
            
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159788)

    End Select

    Exit Sub

End Sub

Private Sub Periodicidade_click()

    iAlterado = REGISTRO_ALTERADO
    Call HabilitaCampos
    
End Sub

Private Sub SpinDataFim_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinDataFim_UpClick

    DataFimExercicio.SetFocus

    If Len(Trim(DataFimExercicio.ClipText)) > 0 Then

        sData = DataFimExercicio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 13657

        DataFimExercicio.PromptInclude = False
        DataFimExercicio.Text = sData
        DataFimExercicio.PromptInclude = True

    End If

    Exit Sub

Erro_SpinDataFim_UpClick:

    Select Case Err

        Case 13657

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159789)

    End Select

    Exit Sub

End Sub

Private Sub SpinDataFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinDataFim_DownClick

    DataFimExercicio.SetFocus

    If Len(Trim(DataFimExercicio.ClipText)) > 0 Then

        sData = DataFimExercicio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 13658

        DataFimExercicio.PromptInclude = False
        DataFimExercicio.Text = sData
        DataFimExercicio.PromptInclude = True
    End If

    Exit Sub

Erro_SpinDataFim_DownClick:

    Select Case Err

        Case 13658

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159790)

    End Select

    Exit Sub

End Sub

Private Sub SpinDataInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinDataInicio_UpClick

    DataInicioExercicio.SetFocus

    If Len(Trim(DataInicioExercicio.ClipText)) > 0 Then

        sData = DataInicioExercicio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 13659

        DataInicioExercicio.Text = sData
    End If

    Exit Sub

Erro_SpinDataInicio_UpClick:

    Select Case Err

        Case 13659

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159791)

    End Select

    Exit Sub

End Sub

Private Sub SpinDataInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinDataInicio_DownClick

    DataInicioExercicio.SetFocus

    If Len(Trim(DataInicioExercicio.ClipText)) > 0 Then

        sData = DataInicioExercicio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 13660

        DataInicioExercicio.Text = sData
    End If

    Exit Sub

Erro_SpinDataInicio_DownClick:

    Select Case Err

        Case 13660

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159792)

    End Select

    Exit Sub

End Sub

Private Sub SpinNumPeriodos_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SpinNumPeriodos_DownClick()

Dim iNumPeriodos As Integer
Dim lErro As Long

On Error GoTo Erro_SpinNumPeriodos_DownClick

    NumPeriodos.SetFocus

    If Len(Trim(NumPeriodos.Text)) = 0 Then
        iNumPeriodos = 0
    Else
        iNumPeriodos = CInt(NumPeriodos.Text)
    End If

    'número de períodos não pode ser menor que 1
    If iNumPeriodos = 1 Then Error 13661

    NumPeriodos.PromptInclude = False
    NumPeriodos.Text = CStr(iNumPeriodos - 1)
    NumPeriodos.PromptInclude = True

    Exit Sub

Erro_SpinNumPeriodos_DownClick:

    Select Case Err

        Case 13661

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159793)

    End Select

    Exit Sub

End Sub

Private Sub SpinNumPeriodos_UpClick()

Dim iNumPeriodos As Integer
Dim lErro As Long

On Error GoTo Erro_SpinNumPeriodos_UpClick

    NumPeriodos.SetFocus

    If Len(Trim(NumPeriodos.Text)) = 0 Then
        iNumPeriodos = 0
    Else
        iNumPeriodos = CInt(NumPeriodos.Text)
    End If

    'número de períodos não pode ser maior que NUM_MAX_PERIODOS
    If iNumPeriodos = NUM_MAX_PERIODOS Then Error 13662

    NumPeriodos.PromptInclude = False
    NumPeriodos.Text = CStr(iNumPeriodos + 1)
    NumPeriodos.PromptInclude = True

    Exit Sub

Erro_SpinNumPeriodos_UpClick:

    Select Case Err

        Case 13662

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159794)

    End Select

    Exit Sub

End Sub

Private Function CalculaPeriodos_Livre(dtDataInicio As Date, dtDataFim As Date, iNumPeriodo As Integer, colPeriodos As Collection) As Long
' Calcula os Periodos quando a Periodicidade é livre

Dim iTotalDias As Integer
Dim iDuracaoPeriodo As Integer
Dim iPeriodo As Integer
Dim dtDataIniPer As Date
Dim dtDataFimPer As Date
Dim objPeriodo As ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_CalculaPeriodos_Livre

    'Calcula o total de dias do Exercicio
    iTotalDias = (dtDataFim - dtDataInicio) + 1

    'Verifica se o numero de periodos requeridos é maior do que o total de dias do Exercicio
    If iTotalDias < iNumPeriodo Then Error 11371

    iDuracaoPeriodo = iTotalDias \ iNumPeriodo

    dtDataIniPer = dtDataInicio
    dtDataFimPer = dtDataInicio + iDuracaoPeriodo - 1

    For iPeriodo = 1 To iNumPeriodo

        Set objPeriodo = New ClassPeriodo

        objPeriodo.dtDataInicio = dtDataIniPer
        objPeriodo.dtDataFim = dtDataFimPer
        objPeriodo.sNomeExterno = "Periodo " & CStr(iPeriodo)

        colPeriodos.Add objPeriodo

        dtDataIniPer = dtDataFimPer + 1
        dtDataFimPer = dtDataIniPer + iDuracaoPeriodo - 1

    Next

    If iTotalDias Mod iNumPeriodo > 0 Then
        lErro = Rotina_Aviso(vbOKOnly, "AVISO_ULTIMO_PERIODO_MAIOR")
    End If

    CalculaPeriodos_Livre = SUCESSO

    Exit Function

Erro_CalculaPeriodos_Livre:

    CalculaPeriodos_Livre = Err

    Select Case Err

        Case 11371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_PERIODOS_MAIOR_TOTAL_DIAS", Err, iNumPeriodo, iTotalDias)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159795)

    End Select

    Exit Function

End Function

Private Function Criar_Exercicio() As Long

Dim lErro As Long
Dim iUltimoExercicio
Dim objExercicio As New ClassExercicio
Dim dtData As Date
Dim iIndice As Integer

On Error GoTo Erro_Criar_Exercicio

    DataFimExercicio.Enabled = True
    SpinDataFim.Enabled = True
    Periodicidade.Enabled = True
    
    'Colocar a periodicidade = MENSAL  --------------- LIVRE
    For iIndice = 0 To Periodicidade.ListCount - 1
        If Periodicidade.ItemData(iIndice) = PERIODICIDADE_MENSAL Then
            Periodicidade.ListIndex = iIndice
            Exit For
        End If
    Next

    NumPeriodos.Enabled = True
    SpinNumPeriodos.Enabled = True
    BotaoGeraPeriodos.Enabled = True
    DataInicioPeriodo.Enabled = True
    objGrid1.iProibidoIncluir = 0
    objGrid1.iProibidoExcluir = 0

    'status = aberto
    Status.Caption = EXERCICIO_DESC_ABERTO

    NumPeriodos.PromptInclude = False
    NumPeriodos.Text = "1"
    NumPeriodos.PromptInclude = True

    NomeExterno.Text = ""

    'Se tiver algum exercicio cadastrado
    If Exercicio.ListCount - 2 >= 0 Then

        iUltimoExercicio = Exercicio.List(Exercicio.ListCount - 2)

        lErro = CF("Exercicio_Le", iUltimoExercicio, objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 13620
    
        'se o exercicio não está cadastrado
        If lErro = 10083 Then Error 10091

        'a data de inicio do proximo exercicio = data fim do ultimo exercicio cadastrado + 1
        dtData = objExercicio.dtDataFim + 1
        
        DataInicioExercicio.Enabled = False
        SpinDataInicio.Enabled = False
        
        
    Else
    
        'se não tem nenhum exercicio cadastrada, sua data inicial será 01/01 do ano corrente.
        dtData = CDate("01/01/" & Year(gdtDataHoje))
        
        DataInicioExercicio.Enabled = True
        SpinDataInicio.Enabled = True
        
    End If
    
    DataInicioExercicio.Text = Format(dtData, "dd/mm/yy")

    'data fim = 31/12 do ano da data início
    DataFimExercicio.Text = "31/12/" & Format(dtData, "yy")

    Call Grid_Limpa(objGrid1)
    
    Call HabilitaCampos

    Criar_Exercicio = SUCESSO

    Exit Function

Erro_Criar_Exercicio:

    Criar_Exercicio = Err
    
    Select Case Err

        Case 10091
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, iUltimoExercicio)

        Case 13620

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159796)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Nome(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Nome

    Set objGridInt.objControle = NomePeriodo

    'testa se o nome do periodo está preenchido
    If Len(Trim(NomePeriodo.Text)) = 0 And GridPeriodos.Row - GridPeriodos.FixedRows < objGridInt.iLinhasExistentes Then Error 13643

    If Len(Trim(NomePeriodo.Text)) > 0 And GridPeriodos.Row - GridPeriodos.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    'verifica se já não existe outro período no exercício com o mesmo nome
    For iIndice = 1 To objGridInt.iLinhasExistentes

        If iIndice <> GridPeriodos.Row Then
            If Trim(NomePeriodo.Text) = Trim(GridPeriodos.TextMatrix(iIndice, GRID_NOME_COL)) Then Error 13644
        End If

    Next
    'critica da coluna 1 fim

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 13645

    Saida_Celula_Nome = SUCESSO

    Exit Function

Erro_Saida_Celula_Nome:

    Saida_Celula_Nome = Err

    Select Case Err

        Case 13643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_PERIODO_VAZIO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 13644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_PERIODO_JA_USADO", Err, Trim(NomePeriodo.Text), iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 13645
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159797)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataIni(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataIni

    Set objGridInt.objControle = DataInicioPeriodo
    
    If Len(DataInicioPeriodo.ClipText) > 0 Then

        lErro = Critica_Campo_DataInicioPeriodo(DataInicioPeriodo.Text)
        If lErro <> SUCESSO Then Error 13647
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 13646

    Saida_Celula_DataIni = SUCESSO

    Exit Function

Erro_Saida_Celula_DataIni:

    Saida_Celula_DataIni = Err

    Select Case Err


        Case 13646, 13647
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159798)

    End Select

    Exit Function

End Function

Private Sub Calcula_Periodos(dtDataInicio As Date, dtDataFim As Date, iDuracaoPeriodo As Integer, colPeriodos As Collection)

Dim iDiaFinal As Integer
Dim iMesFinal As Integer
Dim iAnoFinal As Integer
Dim iTotPeriodo As Integer
Dim objPeriodo As ClassPeriodo
Dim lErro As Long
Dim iNumPeriodo As Integer
Dim dtDataFinal As Date

On Error GoTo Erro_CalCula_Periodos

    iTotPeriodo = 1
    iNumPeriodo = 1
    iMesFinal = Month(dtDataInicio)
    iAnoFinal = Year(dtDataInicio)
    
    Do While iMesFinal > iTotPeriodo * iDuracaoPeriodo
        iTotPeriodo = iTotPeriodo + 1
    Loop
    
    iMesFinal = iTotPeriodo * iDuracaoPeriodo
    
    iDiaFinal = Dias_Mes(iMesFinal, iAnoFinal)
    
    dtDataFinal = CDate(CStr(iDiaFinal) & "/" & CStr(iMesFinal) & "/" & CStr(iAnoFinal))

    If dtDataFim < dtDataFinal Then dtDataFinal = dtDataFim
    
    'adiciona um período ( nome, data inicial e final ) à coleção
    Set objPeriodo = New ClassPeriodo
    objPeriodo.sNomeExterno = "Periodo " & CStr(iNumPeriodo)
    objPeriodo.dtDataInicio = dtDataInicio
    objPeriodo.dtDataFim = dtDataFinal
    objPeriodo.iFechado = PERIODO_ABERTO
    objPeriodo.iFechadoCTB = PERIODO_ABERTO
    colPeriodos.Add objPeriodo
    
    Do While dtDataFim > dtDataFinal And iNumPeriodo < NUM_MAX_PERIODOS
    
        iNumPeriodo = iNumPeriodo + 1
    
        dtDataInicio = dtDataFinal + 1
        
        iMesFinal = Month(dtDataInicio)
        iAnoFinal = Year(dtDataInicio)
    
        iTotPeriodo = 1
    
        Do While iMesFinal > iTotPeriodo * iDuracaoPeriodo
            iTotPeriodo = iTotPeriodo + 1
        Loop

        iMesFinal = iTotPeriodo * iDuracaoPeriodo
        
        iDiaFinal = Dias_Mes(iMesFinal, iAnoFinal)
    
        dtDataFinal = CDate(CStr(iDiaFinal) & "/" & CStr(iMesFinal) & "/" & CStr(iAnoFinal))

        If dtDataFim < dtDataFinal Or iNumPeriodo = NUM_MAX_PERIODOS Then dtDataFinal = dtDataFim
    
        'adiciona um período ( nome, data inicial e final ) à coleção
        Set objPeriodo = New ClassPeriodo
        objPeriodo.sNomeExterno = "Periodo " & CStr(iNumPeriodo)
        objPeriodo.dtDataInicio = dtDataInicio
        objPeriodo.dtDataFim = dtDataFinal
        objPeriodo.iFechado = PERIODO_ABERTO
        objPeriodo.iFechadoCTB = PERIODO_ABERTO
        colPeriodos.Add objPeriodo

    Loop

    Exit Sub

Erro_CalCula_Periodos:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159799)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EXERCICIO
    Set Form_Load_Ocx = Me
    Caption = "Exercicio"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExercicioTela"
    
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


Private Sub LabelNumPeriodos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumPeriodos, Source, X, Y)
End Sub

Private Sub LabelNumPeriodos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumPeriodos, Button, Shift, X, Y)
End Sub

Private Sub LabelPeriodicidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPeriodicidade, Source, X, Y)
End Sub

Private Sub LabelPeriodicidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPeriodicidade, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Status_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Status, Source, X, Y)
End Sub

Private Sub Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Status, Button, Shift, X, Y)
End Sub

Private Sub LabelDataInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataInicio, Source, X, Y)
End Sub

Private Sub LabelDataInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataInicio, Button, Shift, X, Y)
End Sub

Private Sub LabelDataFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataFim, Source, X, Y)
End Sub

Private Sub LabelDataFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataFim, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

