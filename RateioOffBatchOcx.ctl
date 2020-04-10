VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.UserControl RateioOffBatchOcx 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   7215
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   4650
      ScaleHeight     =   690
      ScaleWidth      =   2340
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2400
      Begin VB.CommandButton BotaoFechar 
         Height          =   525
         Left            =   1830
         Picture         =   "RateioOffBatchOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   510
         Left            =   1320
         Picture         =   "RateioOffBatchOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoApurar 
         Height          =   510
         Left            =   105
         Picture         =   "RateioOffBatchOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   1110
      End
   End
   Begin VB.ListBox RateioLista 
      Columns         =   2
      Height          =   1635
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   3180
      Width           =   5085
   End
   Begin VB.TextBox Historico 
      Height          =   345
      Left            =   1095
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2535
      Width           =   3465
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data de Contabilização"
      Height          =   675
      Left            =   150
      TabIndex        =   14
      Top             =   915
      Width           =   6900
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1890
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   735
         TabIndex        =   1
         Top             =   255
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
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
      Begin VB.Label Label4 
         Caption         =   "Data:"
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
         Left            =   165
         TabIndex        =   15
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4665
         TabIndex        =   16
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Periodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5460
         TabIndex        =   17
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Exercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3300
         TabIndex        =   18
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "Exercício:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2370
         TabIndex        =   19
         Top             =   300
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rateios do Tipo Períodos Acumulados - Períodos que serão Rateados"
      Height          =   720
      Left            =   150
      TabIndex        =   13
      Top             =   1680
      Width           =   6900
      Begin VB.ComboBox PeriodoInicial 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   1590
      End
      Begin VB.ComboBox PeriodoFinal 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5100
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Período Inicial:"
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
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Período Final:"
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
         Left            =   3810
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todos"
      Height          =   570
      Left            =   5535
      Picture         =   "RateioOffBatchOcx.ctx":1F72
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4065
      Width           =   1425
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todos"
      Height          =   570
      Left            =   5520
      Picture         =   "RateioOffBatchOcx.ctx":3154
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3270
      Width           =   1425
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   420
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rateios"
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
      TabIndex        =   22
      Top             =   2940
      Width           =   660
   End
   Begin VB.Label LabelLote 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
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
      Left            =   615
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   23
      Top             =   450
      Width           =   450
   End
   Begin VB.Label LabelHistorico 
      AutoSize        =   -1  'True
      Caption         =   "Histórico:"
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
      Left            =   225
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   24
      Top             =   2580
      Width           =   855
   End
End
Attribute VB_Name = "RateioOffBatchOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iExercicioAnterior As Integer

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoHistorico As AdmEvento
Attribute objEventoHistorico.VB_VarHelpID = -1

Public Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166081)
            
    End Select
    
    Exit Function
    
End Function

Private Function Traz_Cabecalho_Tela(dtData As Date, iExercicio As Integer) As Long
'Obtém Periodo e Exercicio correspondentes à data
   
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim objPeriodosFilial As New ClassPeriodosFilial
Dim lErro As Long

On Error GoTo Erro_Traz_Cabecalho_Tela

    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 41463

    'Salva exercício em iExercicio
    iExercicio = objPeriodo.iExercicio
    
    'le exercício
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 41464

    'Exercício não cadastrado
    If lErro = 10083 Then Error 41465

    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
    objPeriodosFilial.sOrigem = MODULO_CONTABILIDADE

    'Le períodos
    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 41467

    'Verifica se periodo está fechado
    If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 41468

    Periodo.Caption = objPeriodo.sNomeExterno

    iExercicioAnterior = objExercicio.iExercicio
    
    Exercicio.Caption = objExercicio.sNomeExterno

    Traz_Cabecalho_Tela = SUCESSO

    Exit Function

Erro_Traz_Cabecalho_Tela:

    Traz_Cabecalho_Tela = Err

    Select Case Err

        Case 41463, 41464, 41467

        Case 41465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 41468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166082)

    End Select

    Exit Function

End Function

Private Function RateioLista_Carga() As Long
'Le rateios off distintos e carrega na lista

Dim lErro As Long
Dim objRateioOff As ClassRateioOff
Dim colRateioOff As New Collection
Dim sNewEntry As String

On Error GoTo Erro_RateioLista_Carga

    'le todos os rateios distintos
    lErro = CF("RateioOff_Le_TodosDistintos", colRateioOff)
    If lErro <> SUCESSO Then Error 41469

    RateioLista.Clear
    
    'Carrega rateios ( codigo + descricao ) na lista
    For Each objRateioOff In colRateioOff

        sNewEntry = CStr(objRateioOff.lCodigo) & " - " & objRateioOff.sDescricao
        RateioLista.AddItem sNewEntry
        RateioLista.ItemData(RateioLista.NewIndex) = objRateioOff.lCodigo

    Next

    RateioLista_Carga = SUCESSO

    Exit Function

Erro_RateioLista_Carga:

    RateioLista_Carga = Err

    Select Case Err

        Case 41469

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166083)

    End Select

    Exit Function

End Function

Private Function ComboPeriodos_Carga(iExercicio As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colPeriodos As New Collection

On Error GoTo Erro_ComboPeriodos_Carga

    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 41539

    PeriodoInicial.Clear
    PeriodoFinal.Clear
        
    For iIndice = 1 To colPeriodos.Count

        PeriodoInicial.AddItem colPeriodos.Item(iIndice).sNomeExterno
        PeriodoInicial.ItemData(PeriodoInicial.NewIndex) = colPeriodos.Item(iIndice).iPeriodo

        PeriodoFinal.AddItem colPeriodos.Item(iIndice).sNomeExterno
        PeriodoFinal.ItemData(PeriodoFinal.NewIndex) = colPeriodos.Item(iIndice).iPeriodo

    Next
    
    PeriodoInicial.ListIndex = 0
    PeriodoFinal.ListIndex = 0

    ComboPeriodos_Carga = SUCESSO
    
    Exit Function

Erro_ComboPeriodos_Carga:

    ComboPeriodos_Carga = Err
    
    Select Case Err
    
        Case 41539
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166084)
            
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim iExercicio As Integer

On Error GoTo Erro_Form_Load

'Inicializar a Data com a data atual
'Inicializar o PeriodoInicial e PeriodoFinal com todos os periodos do exercicio da data atual. Selecionar o primeiro periodo de cada combo.(ListIndex = 0)

    Set objEventoLote = New AdmEvento
    Set objEventoHistorico = New AdmEvento

    Data.Text = Format(gdtDataAtual, "dd/mm/yy")

   'Obtém Periodo e Exercicio correspondentes à data
    lErro = Traz_Cabecalho_Tela(gdtDataAtual, iExercicio)
    If lErro <> SUCESSO Then Error 41470

    lErro = ComboPeriodos_Carga(iExercicio)
    If lErro <> SUCESSO Then Error 41471
    
    lErro = RateioLista_Carga()
    If lErro <> SUCESSO Then Error 41472

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 41470, 41471, 41472

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166085)

    End Select
    
    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data)

End Sub

Private Sub LabelLote_Click()
'Chama o browser de lotes pendentes quando clicar no labelLote

Dim objLote As New ClassLote
Dim dtData As Date
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelLote_Click

    'Obtém Periodo e Exercicio correspondentes à data
    If Len(Data.ClipText) > 0 Then
        dtData = CDate(Data.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 41473

    Else
        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0
    End If

    If Len(Lote.Text) = 0 Then
        objLote.iLote = 0
    Else
        objLote.iLote = CInt(Lote.Text)
    End If

    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo

    colSelecao.Add "CTB"
    colSelecao.Add giFilialEmpresa
    colSelecao.Add 0

    Call Chama_Tela("LotePendenteLista", colSelecao, objLote, objEventoLote)

    Exit Sub

Erro_LabelLote_Click:

    Select Case Err

        Case 41473

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166086)

    End Select

    Exit Sub

End Sub

Private Sub Lote_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Lote)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As ClassLote
Dim iIndice As Integer
Dim sDescricao As String

On Error GoTo Erro_objEventoLote_evSelecao

    Set objLote = obj1

    'Se estiver com a data preenchida ==> verificar se a data está dentro do periodo do lote
    If Len(Data.ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 41474

        'se o periodo/exercicio não correspondem ao periodo/exercicio do lote ==> troca a data
        If objPeriodo.iExercicio <> objLote.iExercicio Or objPeriodo.iPeriodo <> objLote.iPeriodo Then

            'move a data inicial do lote, exercicio e periodo para a tela
            lErro = Move_Data_Tela(objLote)
            If lErro <> SUCESSO Then Error 41475

        End If

    Else

        'se não estiver com a data preenchida
        'move a data inicial do lote, exercicio e periodo para a tela
        lErro = Move_Data_Tela(objLote)
        If lErro <> SUCESSO Then Error 41476

    End If

    Lote.Text = CStr(objLote.iLote)

    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case Err

        Case 41474, 41475, 41476

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166087)

    End Select

    Exit Sub

End Sub

Private Function Move_Data_Tela(objLote As ClassLote) As Long
'Move a data inicial do lote, exercicio e periodo para a tela

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Move_Data_Tela

    lErro = CF("Periodo_Le_ExercicioPeriodo", objLote.iExercicio, objLote.iPeriodo, objPeriodo)
    If lErro <> SUCESSO Then Error 55825

    'Le exercicio
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 41477

    'se o exercicio não estiver cadastrado
    If lErro = 10083 Then Error 41478

    'Verifica se Exercicio está fechado
    If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 41479

    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
    objPeriodosFilial.sOrigem = MODULO_CONTABILIDADE
    

    'Verifica se o Periodo está fechado
    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 41480

    If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 41481

    Data.Text = Format(objPeriodo.dtDataInicio, "dd/mm/yy")

    Periodo.Caption = objPeriodo.sNomeExterno

    iExercicioAnterior = objExercicio.iExercicio
    
    Exercicio.Caption = objExercicio.sNomeExterno

    Move_Data_Tela = SUCESSO

    Exit Function

Erro_Move_Data_Tela:

    Move_Data_Tela = Err

    Select Case Err

        Case 41477, 41480, 55825

        Case 41478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 41479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)

        Case 41481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166088)

    End Select

    Exit Function

End Function

Private Sub LabelHistorico_Click()
'Chama o browser de historicos padrão quando clicar no label historico

Dim objHistPadrao As New ClassHistPadrao
Dim colSelecao As Collection

    If Len(Trim(Historico.Text)) = 0 Then
        objHistPadrao.sDescHistPadrao = ""
    Else
        objHistPadrao.sDescHistPadrao = Historico.Text
    End If

    Call Chama_Tela("HistPadraoLista", colSelecao, objHistPadrao, objEventoHistorico)

End Sub

Private Sub objEventoHistorico_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao

On Error GoTo Erro_objEventoHistorico_evSelecao

    Set objHistPadrao = obj1

    Historico.Text = objHistPadrao.sDescHistPadrao

    Me.Show

    Exit Sub

Erro_objEventoHistorico_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166089)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Tela() As Long

Dim lErro As Long
Dim iExercicio As Integer

On Error GoTo Erro_Inicializa_Tela

    Call Limpa_Tela(Me)

    Data.Text = Format(gdtDataAtual, "dd/mm/yy")

    lErro = Traz_Cabecalho_Tela(gdtDataAtual, iExercicio)
    If lErro <> SUCESSO Then Error 41483
    
    lErro = ComboPeriodos_Carga(iExercicio)
    If lErro <> SUCESSO Then Error 41540

    lErro = RateioLista_Carga()
    If lErro <> SUCESSO Then Error 41536

    Call BotaoDesmarcarTodos_Click

    Inicializa_Tela = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tela:

    Inicializa_Tela = Err
    
    Select Case Err
    
        Case 41483
            Data.SetFocus
            
        Case 41536, 41540
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166090)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lDoc As Long
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Inicializa_Tela()
    If lErro <> SUCESSO Then Error 41531
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 41482, 41531

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166091)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoLote = Nothing
    Set objEventoHistorico = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166092)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer
    
    For iIndice = 0 To RateioLista.ListCount - 1

        RateioLista.Selected(iIndice) = True

    Next

    RateioLista.ListIndex = 0
    
    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click()

Dim iIndice As Integer

    For iIndice = 0 To RateioLista.ListCount - 1

        RateioLista.Selected(iIndice) = False

    Next
    
    RateioLista.ListIndex = 0

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim objLote As New ClassLote
Dim vbMsgRes As VbMsgBoxResult
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) > 0 Then

        'critica a data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 41484

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)
        
        'Le os periodos referentes a dtData
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 41485

        'Verifica se Exercicio está fechado
        lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 41486

        'Exercicio não cadastrado
        If lErro = 10083 Then Error 41487

        'Verifica se o exercício está fechado
        If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 41488

        objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
        objPeriodosFilial.iExercicio = objPeriodo.iExercicio
        objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
        objPeriodosFilial.sOrigem = MODULO_CONTABILIDADE

        'le periodos
        lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
        If lErro <> SUCESSO Then Error 41489

        'Verifica se o periodo está fechado
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 41490

        If Len(Lote.Text) > 0 Then

            objLote.iLote = CInt(Lote.Text)

            objLote.iFilialEmpresa = giFilialEmpresa
            objLote.sOrigem = MODULO_CONTABILIDADE
            objLote.iExercicio = objPeriodo.iExercicio
            objLote.iPeriodo = objPeriodo.iPeriodo

            'verifica se o lote  está atualizado
            lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
            If lErro <> SUCESSO Then Error 41491

            'Se é um lote que já foi contabilizado, não pode sofrer alteração
            If iLoteAtualizado = LOTE_ATUALIZADO Then Error 41492

            'Le a tabela lote pendente
            lErro = CF("LotePendente_Le", objLote)
            If lErro <> SUCESSO And lErro <> 5435 Then Error 41493

            'Se o lote não está cadastrado
            If lErro = 5435 Then Error 41494

            'checa se o lote pertence ao periodo em questão
            If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 41495

        End If

        'Preenche campo de periodo
        Periodo.Caption = objPeriodo.sNomeExterno
        Exercicio.Caption = objExercicio.sNomeExterno
        
        If (iExercicioAnterior <> objExercicio.iExercicio) Then
            lErro = ComboPeriodos_Carga(objExercicio.iExercicio)
            If lErro <> SUCESSO Then Error 41541
        End If

        iExercicioAnterior = objExercicio.iExercicio
        
    Else

        Periodo.Caption = ""
        Exercicio.Caption = ""
                    
        PeriodoInicial.Clear
        PeriodoFinal.Clear

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True
    
    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err

        Case 41484, 41485, 41486, 41489, 41491, 41493, 41541
            
        Case 41487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)
    
        Case 41488
            'Não é possível fazer lançamentos em exercício fechado
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)
    
        Case 41490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
        
        Case 41492
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objLote.iExercicio, objLote.iPeriodo, MODULO_CONTABILIDADE)
        
        Case 41494
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_INEXISTENTE", Err, objLote.iFilialEmpresa, objLote.iLote, MODULO_CONTABILIDADE, objLote.iPeriodo, objLote.iExercicio)
            If vbMsgRes = vbYes Then
                'Se respondeu que deseja criar LOTE
                Call Chama_Tela("LoteTela", objLote)
            End If

        Case 41495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166093)

    End Select

    Exit Sub

End Sub

Private Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As New ClassLote
Dim sNomeExterno As String
Dim objExercicio As New ClassExercicio
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection

On Error GoTo Erro_Lote_Validate

    If Len(Trim(Lote.Text)) > 0 And Len(Trim(Data.ClipText)) > 0 Then

        objLote.iLote = CInt(Lote.Text)
        objLote.iFilialEmpresa = giFilialEmpresa
        objLote.sOrigem = MODULO_CONTABILIDADE

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 41496

        objLote.iExercicio = objPeriodo.iExercicio
        objLote.iPeriodo = objPeriodo.iPeriodo

        'verifica se o lote  está atualizado
        lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
        If lErro <> SUCESSO Then Error 41497

        'Se é um lote que já foi contabilizado, não pode sofrer alteração
        If iLoteAtualizado = LOTE_ATUALIZADO Then Error 41498

        lErro = CF("LotePendente_Le", objLote)
        If lErro <> SUCESSO And lErro <> 5435 Then Error 41499

        'Se o lote não está cadastrado
        If lErro = 5435 Then Error 41500

        If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 41501

    End If

    Exit Sub

Erro_Lote_Validate:

    Cancel = True
    
    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Me.Show
    End If

    Select Case Err

        Case 41496, 41497, 41499

        Case 41498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objPeriodo.iExercicio, objPeriodo.iPeriodo, MODULO_CONTABILIDADE)

        Case 41500
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_INEXISTENTE", objLote.iFilialEmpresa, objLote.iLote, MODULO_CONTABILIDADE, objPeriodo.iPeriodo, objPeriodo.iExercicio)

            If vbMsgRes = vbYes Then
                'Se respondeu que deseja criar LOTE
                Call Chama_Tela("LoteTela", objLote)
                
            End If
        
        Case 41501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166094)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 41502

        Data.Text = sData

    End If
    
    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 41502

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166095)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 41503

        Data.Text = sData

    End If
    
    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 41503

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166096)

    End Select

    Exit Sub

End Sub

Private Sub BotaoApurar_Click()

Dim lErro As Long
Dim sNomeArqParam As String
Dim objRateioOffBatch As New ClassRateioOffBatch

On Error GoTo Erro_BotaoApurar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos estão preenchidos
    lErro = Verifica_CamposPreenchidos()
    If lErro <> SUCESSO Then Error 41504

    'Critica dados dos campos
    lErro = Critica_Campos()
    If lErro <> SUCESSO Then Error 41505

    'obtem dados da tela
    lErro = Move_Tela_Memoria(objRateioOffBatch)
    If lErro <> SUCESSO Then Error 41506

    'preparo da rotina batch
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 41507

    'chamada da rotina batch
    lErro = CF("Rotina_RateioOff", sNomeArqParam, objRateioOffBatch.iLote, objRateioOffBatch.dtData, objRateioOffBatch.iPeriodoInicial, objRateioOffBatch.iPeriodoFinal, objRateioOffBatch.colRateios, objRateioOffBatch.iFilialEmpresa, objRateioOffBatch.sHistorico)
    If lErro <> SUCESSO Then Error 41508

    'Reinicializa a Tela
    lErro = Inicializa_Tela()
    If lErro <> SUCESSO Then Error 41532
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoApurar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case Err

        Case 41504 To 41508, 41532

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166097)

    End Select

    Exit Sub

End Sub

Private Function Verifica_CamposPreenchidos() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iSelecionado As Integer

On Error GoTo Erro_Verifica_CamposPreenchidos

    iSelecionado = 0

    If Len(Trim(Lote.ClipText)) = 0 Then Error 41509

    If Len(Trim(Data.ClipText)) = 0 Then Error 41510

    If PeriodoInicial.ListIndex = -1 Then Error 41511

    If PeriodoFinal.ListIndex = -1 Then Error 41512

    If RateioLista.ListCount = 0 Then Error 41514

    For iIndice = 0 To RateioLista.ListCount - 1
        If RateioLista.Selected(iIndice) = True Then
            iSelecionado = 1
            Exit For
        End If
    Next

    If iSelecionado = 0 Then Error 41515

    Verifica_CamposPreenchidos = SUCESSO

    Exit Function

Erro_Verifica_CamposPreenchidos:

    Verifica_CamposPreenchidos = Err

    Select Case Err

        Case 41509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_PREENCHIDO", Err)

        Case 41510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CONTABIL_NAO_PREENCHIDA", Err)

        Case 41511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_NAO_SELECIONADO", Err)

        Case 41512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_FINAL_NAO_SELECIONADO", Err)

        Case 41514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RATEIOS_INEXISTENTES", Err)

        Case 41515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RATEIOS_NAO_INFORMADOS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166098)

    End Select

    Exit Function

End Function

Private Function Critica_Campos() As Long

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim objPeriodosFilial As New ClassPeriodosFilial
Dim objLote As New ClassLote
Dim iLoteAtualizado As Integer

On Error GoTo Erro_Critica_Campos

    'verifica se o periodo final é menor que o inicial
    If PeriodoFinal.ItemData(PeriodoFinal.ListIndex) < PeriodoInicial.ItemData(PeriodoInicial.ListIndex) Then Error 41527

    dtData = CDate(Data.Text)

    'Le o periodo referente a data
    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 41516

    'verifica se o período final maior é maior que o período de contabilização
    If PeriodoFinal.ItemData(PeriodoFinal.ListIndex) > objPeriodo.iPeriodo Then Error 41528

    'Verifica se Exercicio está fechado
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 41517

    'Exercicio não cadastrado
    If lErro = 10083 Then Error 41518

    'verifica se o exercício está fechado
    If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 41519

    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
    objPeriodosFilial.sOrigem = MODULO_CONTABILIDADE

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 41520

    If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 41521

    objLote.iLote = CInt(Lote.Text)

    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = MODULO_CONTABILIDADE
    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo

    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 41522

    'Se é um lote que já foi contabilizado, não pode sofrer alteração
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 41523

    lErro = CF("LotePendente_Le", objLote)
    If lErro <> SUCESSO And lErro <> 5435 Then Error 41524

    'Se o lote não está cadastrado
    If lErro = 5435 Then Error 41525
    
    'checa se o lote pertence ao periodo em questão
    If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 41526

    Critica_Campos = SUCESSO

    Exit Function

Erro_Critica_Campos:

    Critica_Campos = Err

    Select Case Err

        Case 41516, 41517, 41520, 41522, 41524

        Case 41518
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 41519
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)

        Case 41521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
            
        Case 41523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objLote.iExercicio, objLote.iPeriodo, MODULO_CONTABILIDADE)

        Case 41525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, MODULO_CONTABILIDADE, objLote.iPeriodo, objLote.iExercicio, objLote.iLote)

        Case 41526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)

        Case 41527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOFINAL_MENOR", Err)
            
        Case 41528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOFINAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166099)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objRateioOffBatch As ClassRateioOffBatch) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objRateioOffBatch.dtData = CDate(Data.Text)
    objRateioOffBatch.iFilialEmpresa = giFilialEmpresa
    objRateioOffBatch.iLote = CInt(Lote.Text)
    objRateioOffBatch.sHistorico = Historico.Text
    objRateioOffBatch.iPeriodoInicial = PeriodoInicial.ItemData(PeriodoInicial.ListIndex)
    objRateioOffBatch.iPeriodoFinal = PeriodoFinal.ItemData(PeriodoFinal.ListIndex)
    
    Set objRateioOffBatch.colRateios = New Collection
        
    For iIndice = 0 To RateioLista.ListCount - 1
        If RateioLista.Selected(iIndice) = True Then
            objRateioOffBatch.colRateios.Add RateioLista.ItemData(iIndice)
        End If
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166100)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RATEIO_OFF_LINE_PROCESSAM_BATCH
    Set Form_Load_Ocx = Me
    Caption = "Rateio Off-Line - Processamento Batch"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RateioOffBatch"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Lote Then
            Call LabelLote_Click
        ElseIf Me.ActiveControl Is Historico Then
            Call LabelHistorico_Click
        End If
    
    End If

End Sub




Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Periodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Periodo, Source, X, Y)
End Sub

Private Sub Periodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Periodo, Button, Shift, X, Y)
End Sub

Private Sub Exercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Exercicio, Source, X, Y)
End Sub

Private Sub Exercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Exercicio, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelLote, Source, X, Y)
End Sub

Private Sub LabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelLote, Button, Shift, X, Y)
End Sub

Private Sub LabelHistorico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistorico, Source, X, Y)
End Sub

Private Sub LabelHistorico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistorico, Button, Shift, X, Y)
End Sub

