VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.UserControl LoteEstorno1Ocx 
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   5985
   Begin VB.Frame Frame1 
      Caption         =   "Lote de Estorno"
      Height          =   1845
      Left            =   120
      TabIndex        =   5
      Top             =   1755
      Width           =   5670
      Begin MSMask.MaskEdBox LoteEstorno 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   375
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
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
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   2235
         TabIndex        =   6
         Top             =   900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEstorno 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   915
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
      Begin VB.Label OrigemEstorno 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3780
         TabIndex        =   7
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
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
         Left            =   3075
         TabIndex        =   8
         Top             =   480
         Width           =   660
      End
      Begin VB.Label LabelLoteEstorno 
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
         Left            =   570
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   420
         Visible         =   0   'False
         Width           =   450
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
         Left            =   2865
         TabIndex        =   10
         Top             =   915
         Width           =   870
      End
      Begin VB.Label ExercicioEstorno 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3780
         TabIndex        =   11
         Top             =   900
         Width           =   1530
      End
      Begin VB.Label PeriodoEstorno 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3780
         TabIndex        =   12
         Top             =   1380
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   540
         TabIndex        =   13
         Top             =   945
         Width           =   480
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
         Left            =   3000
         TabIndex        =   14
         Top             =   1410
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lote a ser Estornado"
      Height          =   1530
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5670
      Begin VB.Label Label3 
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
         Left            =   3000
         TabIndex        =   15
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Periodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3780
         TabIndex        =   16
         Top             =   975
         Width           =   1530
      End
      Begin VB.Label Exercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1125
         TabIndex        =   17
         Top             =   975
         Width           =   1530
      End
      Begin VB.Label Label9 
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
         Left            =   210
         TabIndex        =   18
         Top             =   1005
         Width           =   870
      End
      Begin VB.Label Label10 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   630
         TabIndex        =   19
         Top             =   465
         Width           =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
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
         Left            =   3075
         TabIndex        =   20
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Origem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3780
         TabIndex        =   21
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label Lote 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1125
         TabIndex        =   22
         Top             =   450
         Width           =   705
      End
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3255
      Picture         =   "LoteEstorno1Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3825
      Width           =   975
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1305
      Picture         =   "LoteEstorno1Ocx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3825
      Width           =   975
   End
End
Attribute VB_Name = "LoteEstorno1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private objLancamento_Cabecalho1 As New ClassLancamento_Cabecalho
Private objLote1 As New ClassLote
Private objBrowseConfigura1 As AdmBrowseConfigura
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Dim iLote_Lost_Focus As Integer
Dim iData_Lost_Focus As Integer

Private Sub BotaoCancelar_Click()

    objBrowseConfigura1.iTelaOK = CANCELA

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a Data dos Lançamentos de Estorno está preenchida
    If Len(DataEstorno.ClipText) = 0 Then Error 36851

'    'Verifica se o Lote de Estorno está preenchido
'    If Len(LoteEstorno.ClipText) = 0 Then Error 36852

    'Preenche Objeto Lançamento_Cabeçalho
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)
'    objLancamento_Cabecalho.iLote = CInt(LoteEstorno.ClipText)
    objLancamento_Cabecalho.dtData = CDate(DataEstorno.Text)


    'grava o estorno. objLancamento_Cabecalho1 contém o lote a ser extornado. objLancamento_Cabecalho contém algumas informacoes dos lançamentos de estorno a serem criados
    lErro = CF("Lancamento_Grava_Estorno", objLancamento_Cabecalho, objLancamento_Cabecalho1)
    If lErro <> SUCESSO Then Error 36853

    objBrowseConfigura1.iTelaOK = OK

    GL_objMDIForm.MousePointer = vbDefault
    
    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 36851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", Err)
            DataEstorno.SetFocus

        Case 36852
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_LOTE_NAO_PREENCHIDO", Err)
            LoteEstorno.SetFocus

        Case 36853

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162473)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objLote As ClassLote, objBrowseConfigura As AdmBrowseConfigura) As Long

Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set objBrowseConfigura1 = objBrowseConfigura

    'guarda os dados de objLote para ser usado na gravação do estorno
    objLancamento_Cabecalho1.iFilialEmpresa = objLote.iFilialEmpresa
    objLancamento_Cabecalho1.sOrigem = objLote.sOrigem
    objLancamento_Cabecalho1.iExercicio = objLote.iExercicio
    objLancamento_Cabecalho1.iPeriodoLote = objLote.iPeriodo
    objLancamento_Cabecalho1.iLote = objLote.iLote

    Lote.Caption = objLote.iLote

    lErro = CF("Exercicio_Le", objLote.iExercicio, objExercicio)
    If lErro <> SUCESSO Then Error 36858

    Exercicio.Caption = objExercicio.sNomeExterno

    lErro = CF("Periodo_Le_ExercicioPeriodo", objLote.iExercicio, objLote.iPeriodo, objPeriodo)
    If lErro <> SUCESSO Then Error 36859

    Periodo.Caption = objPeriodo.sNomeExterno

    Origem.Caption = gobjColOrigem.Descricao(objLote.sOrigem)

    OrigemEstorno.Caption = gobjColOrigem.Descricao(objLote.sOrigem)


    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 36858, 36859

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162474)

    End Select

    Exit Function

End Function

Private Sub DataEstorno_GotFocus()
    iData_Lost_Focus = 0
    Call MaskEdBox_TrataGotFocus(DataEstorno)
End Sub

Private Sub DataEstorno_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lDoc As Long
Dim sNomeExterno As String
Dim objLote As New ClassLote
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_DataEstorno_Validate

    If iLote_Lost_Focus = 0 Then

        If Len(DataEstorno.ClipText) > 0 Then

            lErro = Data_Critica(DataEstorno.Text)
            If lErro <> SUCESSO Then Error 36860

            'Obtém Periodo e Exercicio correspondentes à data
            dtData = CDate(DataEstorno.Text)

            lErro = CF("Periodo_Le", dtData, objPeriodo)
            If lErro <> SUCESSO Then Error 36872

            'Verifica se Exercicio está fechado
            lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
            If lErro <> SUCESSO And lErro <> 10083 Then Error 36873

            'Exercicio não cadastrado
            If lErro = 10083 Then Error 36874

            If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 36875

            objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
            objPeriodosFilial.iExercicio = objPeriodo.iExercicio
            objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
            objPeriodosFilial.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)
            

            lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
            If lErro <> SUCESSO Then Error 36876

            If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 36877

            'checa se o lote pertence ao periodo em questão
            If Len(LoteEstorno.Text) > 0 Then

                objLote.iLote = CInt(LoteEstorno.Text)

                objLote.iFilialEmpresa = giFilialEmpresa
                objLote.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)
                objLote.iExercicio = objPeriodo.iExercicio
                objLote.iPeriodo = objPeriodo.iPeriodo

                'verifica se o lote  está atualizado
                lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
                If lErro <> SUCESSO Then Error 36878

                'Se é um lote que já foi contabilizado, não pode sofrer alteração
                If iLoteAtualizado = LOTE_ATUALIZADO Then Error 36879

                lErro = CF("LotePendente_Le", objLote)
                If lErro <> SUCESSO And lErro <> 5435 Then Error 36880

                'Se o lote não está cadastrado
                If lErro = 5435 Then Error 36881

                If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 36882


            End If

            'Preenche campo de periodo
            PeriodoEstorno.Caption = objPeriodo.sNomeExterno

            ExercicioEstorno.Caption = objExercicio.sNomeExterno

        Else

            PeriodoEstorno.Caption = ""

            ExercicioEstorno.Caption = ""

        End If

    End If

    Exit Sub

Erro_DataEstorno_Validate:

    Cancel = True


    Select Case Err

        Case 36860
            iData_Lost_Focus = 1

        Case 36872, 36873, 36876

        Case 36874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 36875
            'Não é possível fazer lançamentos em exercício fechado
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objLote.iExercicio)
            iData_Lost_Focus = 1

        Case 36877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)
            iData_Lost_Focus = 1

        Case 36879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objLote.iExercicio, objLote.iPeriodo, OrigemEstorno.Caption)
            iData_Lost_Focus = 1

        Case 36880, 36878

        Case 36881
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            iData_Lost_Focus = 1

        Case 36882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)
            iData_Lost_Focus = 1

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162475)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoLote = Nothing
        
    Set objLancamento_Cabecalho1 = Nothing
    Set objLote1 = Nothing
    Set objBrowseConfigura1 = Nothing
    
End Sub

Private Sub LoteEstorno_GotFocus()
    iLote_Lost_Focus = 0
    Call MaskEdBox_TrataGotFocus(LoteEstorno)

End Sub

Private Sub LoteEstorno_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As New ClassLote
Dim sNomeExterno As String
Dim objExercicio As New ClassExercicio
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection

On Error GoTo Erro_LoteEstorno_Validate

    If iData_Lost_Focus = 0 Then

        If Len(LoteEstorno.Text) > 0 And Len(DataEstorno.ClipText) > 0 Then

            objLote.iLote = CInt(LoteEstorno.Text)
            objLote.iFilialEmpresa = giFilialEmpresa
            objLote.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)

            'Obtém Periodo e Exercicio correspondentes à data
            dtData = CDate(DataEstorno.Text)

            lErro = CF("Periodo_Le", dtData, objPeriodo)
            If lErro <> SUCESSO Then Error 36883

            objLote.iExercicio = objPeriodo.iExercicio
            objLote.iPeriodo = objPeriodo.iPeriodo

            'verifica se o lote  está atualizado
            lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
            If lErro <> SUCESSO Then Error 36884

            'Se é um lote que já foi contabilizado, não pode sofrer alteração
            If iLoteAtualizado = LOTE_ATUALIZADO Then Error 36885

            lErro = CF("LotePendente_Le", objLote)
            If lErro <> SUCESSO And lErro <> 5435 Then Error 36886

            'Se o lote não está cadastrado
            If lErro = 5435 Then Error 36887

            If giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 36888


        End If

    End If

    Exit Sub

Erro_LoteEstorno_Validate:

    Cancel = True


    Select Case Err

        Case 36883, 36884, 36886

        Case 36885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLote.iFilialEmpresa, objLote.iLote, objPeriodo.iExercicio, objPeriodo.iPeriodo, OrigemEstorno.Caption)
            iLote_Lost_Focus = 1

        Case 36887
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INEXISTENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            iLote_Lost_Focus = 1

        Case 36888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)
            iLote_Lost_Focus = 1


        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162476)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim objExercicio As New ClassExercicio
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Form_Load

    iData_Lost_Focus = 0
    iLote_Lost_Focus = 0
    
    Set objEventoLote = New AdmEvento

    DataEstorno.Text = Format(gdtDataAtual, "dd/mm/yy")

    'le periodos
    lErro = CF("Periodo_Le", gdtDataAtual, objPeriodo)
    If lErro <> SUCESSO Then Error 36857

    'le exercício
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 36856

    'Exercício não cadastrado
    If lErro = 10083 Then Error 41630

    'Verifica se Exercicio está fechado
    If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 41631

'    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
'    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
'    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
'    objPeriodosFilial.sOrigem = MODULO_CONTABILIDADE
'
'    'Le períodos
'    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
'    If lErro <> SUCESSO Then Error 41632

'    'Verifica se periodo está fechado
'    If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 41633

    PeriodoEstorno.Caption = objPeriodo.sNomeExterno
    ExercicioEstorno.Caption = objExercicio.sNomeExterno
    OrigemEstorno.Caption = gobjColOrigem.Descricao(MODULO_CONTABILIDADE)


    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 36856, 36857, 41632

        Case 41630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 41631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)

        Case 41633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162477)

    End Select

    Exit Sub

End Sub

Private Sub LabelLoteEstorno_Click()

Dim objLote As New ClassLote
Dim dtData As Date
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelLoteEstorno_Click

    'Obtém Periodo e Exercicio correspondentes à data
    If Len(DataEstorno.ClipText) > 0 Then
        dtData = CDate(DataEstorno.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36862

    Else
        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0
    End If

    If Len(LoteEstorno.Text) = 0 Then
        objLote.iLote = 0
    Else
        objLote.iLote = CInt(LoteEstorno.Text)
    End If

    objLote.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)
    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo

    Call Chama_Tela_Modal("LotePendenteListaModal", colSelecao, objLote, objEventoLote)

    Exit Sub

Erro_LabelLoteEstorno_Click:

    Select Case Err

        Case 36862

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162478)

    End Select

    Exit Sub

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
    If Len(DataEstorno.ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(DataEstorno.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36863

        'se o periodo/exercicio não corresponde ao periodo/exercicio do lote ==> troca a data
        If objPeriodo.iExercicio <> objLote.iExercicio Or objPeriodo.iPeriodo <> objLote.iPeriodo Then

            'move a data inicial do lote, exercicio e periodo para a tela
            lErro = Move_Data_Tela(objLote)
            If lErro <> SUCESSO Then Error 36864

        End If

    Else

        'se não estiver com a data preenchida
        'move a data inicial do lote, exercicio e periodo para a tela
        lErro = Move_Data_Tela(objLote)
        If lErro <> SUCESSO Then Error 36865

    End If

    LoteEstorno.Text = CStr(objLote.iLote)

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case Err

        Case 36863, 36864, 36865  'Erro já tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162479)

    End Select

    Exit Sub

End Sub

Private Function Move_Data_Tela(objLote As ClassLote) As Long

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Move_Data_Tela

    lErro = CF("Periodo_Le_ExercicioPeriodo", objLote.iExercicio, objLote.iPeriodo, objPeriodo)
    If lErro <> SUCESSO Then Error 36866

    'Verifica se Exercicio está fechado
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 36867

    'se o exercicio não estiver cadastrado
    If lErro = 10083 Then Error 36869

    If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 36870

    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
    objPeriodosFilial.sOrigem = gobjColOrigem.Origem(OrigemEstorno.Caption)


    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 36868

    If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 36871

    DataEstorno.Text = Format(objPeriodo.dtDataInicio, "dd/mm/yy")

    PeriodoEstorno.Caption = objPeriodo.sNomeExterno

    ExercicioEstorno.Caption = objExercicio.sNomeExterno

    Move_Data_Tela = SUCESSO

    Exit Function

Erro_Move_Data_Tela:

    Move_Data_Tela = Err

    Select Case Err

        Case 36866, 36867, 36868

        Case 36869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 36870
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)

        Case 36871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162480)

    End Select

    Exit Function

End Function

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    If Len(Trim(DataEstorno.ClipText)) > 0 Then

        sData = DataEstorno.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 41634

        DataEstorno.Text = sData

    End If

    Call DataEstorno_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 41634

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162481)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    If Len(Trim(DataEstorno.ClipText)) > 0 Then

        sData = DataEstorno.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 41635

        DataEstorno.Text = sData

    End If

    Call DataEstorno_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 41635

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162482)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EXTORNO_LOTE_CONTABILIZADO1
    Set Form_Load_Ocx = Me
    Caption = "Estorno de Lote Contabilizado - Lote de Estorno"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LoteEstorno1"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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
        
        If Me.ActiveControl Is LoteEstorno Then
            Call LabelLoteEstorno_Click
        End If
    
    End If

End Sub


Private Sub OrigemEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OrigemEstorno, Source, X, Y)
End Sub

Private Sub OrigemEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OrigemEstorno, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelLoteEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelLoteEstorno, Source, X, Y)
End Sub

Private Sub LabelLoteEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelLoteEstorno, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub ExercicioEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ExercicioEstorno, Source, X, Y)
End Sub

Private Sub ExercicioEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ExercicioEstorno, Button, Shift, X, Y)
End Sub

Private Sub PeriodoEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PeriodoEstorno, Source, X, Y)
End Sub

Private Sub PeriodoEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PeriodoEstorno, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
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

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Origem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Origem, Source, X, Y)
End Sub

Private Sub Origem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Origem, Button, Shift, X, Y)
End Sub

Private Sub Lote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Lote, Source, X, Y)
End Sub

Private Sub Lote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Lote, Button, Shift, X, Y)
End Sub

