VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ApuraPeriodoOcx 
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   LockControls    =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   7680
   Begin VB.CheckBox ZeraRD 
      Caption         =   "Zera receitas e despesas"
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
      Left            =   1065
      TabIndex        =   5
      Top             =   3255
      Width           =   2595
   End
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   4980
      ScaleHeight     =   690
      ScaleWidth      =   2445
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   150
      Width           =   2505
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   1935
         Picture         =   "ApuraPeriodoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoApurar 
         Height          =   510
         Left            =   120
         Picture         =   "ApuraPeriodoOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   510
         Left            =   1455
         Picture         =   "ApuraPeriodoOcx.ctx":1A40
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   90
         Width           =   390
      End
   End
   Begin VB.ListBox Historicos 
      Height          =   2790
      ItemData        =   "ApuraPeriodoOcx.ctx":1F72
      Left            =   4800
      List            =   "ApuraPeriodoOcx.ctx":1F74
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1170
      Width           =   2700
   End
   Begin VB.ComboBox Exercicio 
      Height          =   315
      ItemData        =   "ApuraPeriodoOcx.ctx":1F76
      Left            =   2295
      List            =   "ApuraPeriodoOcx.ctx":1F78
      OLEDropMode     =   1  'Manual
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   1860
   End
   Begin VB.ComboBox PeriodoInicial 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1860
   End
   Begin VB.ComboBox PeriodoFinal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1545
      Width           =   1860
   End
   Begin VB.TextBox Historico 
      Height          =   330
      Left            =   1065
      MaxLength       =   150
      TabIndex        =   6
      Top             =   3675
      Width           =   3585
   End
   Begin MSMask.MaskEdBox ContaContraPartida 
      Height          =   315
      Left            =   2295
      TabIndex        =   4
      Top             =   2730
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
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
   Begin MSMask.MaskEdBox ContaResultado 
      Height          =   315
      Left            =   2295
      TabIndex        =   3
      Top             =   2130
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
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
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   2790
      Left            =   4800
      TabIndex        =   8
      Top             =   1170
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4921
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1395
      TabIndex        =   13
      Top             =   465
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Left            =   960
      TabIndex        =   14
      Top             =   990
      Width           =   1320
   End
   Begin VB.Label Label11 
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
      Left            =   1080
      TabIndex        =   15
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta de Contra Partida:"
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
      TabIndex        =   16
      Top             =   2790
      Width           =   2130
   End
   Begin VB.Label Label9 
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
      Height          =   210
      Left            =   165
      TabIndex        =   17
      Top             =   3735
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conta de Resultado:"
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
      Left            =   480
      TabIndex        =   18
      Top             =   2190
      Width           =   1770
   End
   Begin VB.Label Hist 
      AutoSize        =   -1  'True
      Caption         =   "Históricos"
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
      Left            =   4845
      TabIndex        =   19
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label LblTvws 
      AutoSize        =   -1  'True
      Caption         =   "Plano de Contas"
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
      Left            =   4845
      TabIndex        =   20
      Top             =   975
      Width           =   1410
   End
End
Attribute VB_Name = "ApuraPeriodoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Const CONTA_CONTRA_PARTIDA = "1"
Const CONTA_RESULTADO = "2"

Private Sub Exercicio_Click()

Dim lErro As Long
Dim iExercicio As Integer

On Error GoTo Erro_Exercicio_Click

    If Exercicio.ListIndex <> -1 Then

        iExercicio = Exercicio.ItemData(Exercicio.ListIndex)

        lErro = Preenche_CombosPeriodos(iExercicio)
        If lErro <> SUCESSO Then Error 11664

    End If

    Exit Sub
    
Erro_Exercicio_Click:

    Select Case Err
    
        Case 11664
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143124)
            
        End Select
        
        Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim iExercicio As Integer
Dim iPosInicial As Integer
Dim iLote As Integer

On Error GoTo Erro_Form_Load

    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 11685

    iPosInicial = -1

    lErro = Preenche_ComboExercicio(iPosInicial)
    If lErro <> SUCESSO Then Error 11666

    If Exercicio.ListCount = 0 Then Error 11667

    Exercicio.ListIndex = 0

    If iPosInicial >= 0 Then
    
        iExercicio = CInt(Exercicio.ItemData(0))

        lErro = Preenche_CombosPeriodos(iExercicio)
        If lErro <> SUCESSO Then Error 11668

    End If
    
    'Inicializa a Lista de Plano de Contas
    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
    If lErro <> SUCESSO Then Error 11669

    TvwContas.Tag = CONTA_RESULTADO
    
    'Inicializa a Lista de Historicos
    lErro = Carga_ListBox_Historico()
    If lErro <> SUCESSO Then Error 11670
   
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 11666, 11668, 11669, 11670, 11685
        
        Case 11667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOS_FECHADOS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143125)

    End Select

    Exit Sub

End Sub

Private Sub BotaoApurar_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iIndex As Integer
Dim iExercicio As Integer
Dim iPeriodoInicial As Integer
Dim iPeriodoFinal As Integer
Dim sConta As String
Dim sContaContraPartida As String
Dim sContaResultado As String
Dim sHistorico As String
Dim iContaPreenchida As Integer
Dim sContaResultadoNivel1 As String
Dim sContaContraPartidaNivel1 As String
Dim colPlanoConta As New Collection
Dim colContaCategoria As New Collection
Dim colContasApuracao As New Collection
Dim objContaCategoria As ClassContaCategoria
Dim objPlanoConta As ClassPlanoConta
Dim sNomeArqParam As String
Dim iZeraRD As Integer

On Error GoTo Erro_BotaoApurar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Exercicio.ListIndex = -1 Then Error 11641

    iExercicio = Exercicio.ItemData(Exercicio.ListIndex)

    If PeriodoInicial.ListIndex = -1 Then Error 11642
    
    iPeriodoInicial = PeriodoInicial.ItemData(PeriodoInicial.ListIndex)

    If PeriodoFinal.ListIndex = -1 Then Error 11643
    
    iPeriodoFinal = PeriodoFinal.ItemData(PeriodoFinal.ListIndex)

    If iPeriodoFinal < iPeriodoInicial Then Error 11644

    If Len(Trim(ContaContraPartida.ClipText)) = 0 And ZeraRD.Value = vbUnchecked Then Error 11645

    If Len(Trim(ContaResultado.ClipText)) = 0 Then Error 11646
    
    If ZeraRD.Value = vbUnchecked Then
        iZeraRD = DESMARCADO
    Else
        iZeraRD = MARCADO
    End If

    sConta = ContaResultado.Text

    'Guarda a conta ja Formatada em sContaResultado
    lErro = CF("Conta_Formata", sConta, sContaResultado, iContaPreenchida)
    If lErro <> SUCESSO Then Error 9798

    sConta = ContaContraPartida.Text

    If Len(Trim(ContaContraPartida.ClipText)) <> 0 Then
        'Guarda a conta ja Formatada em sContaContraPartida
        lErro = CF("Conta_Formata", sConta, sContaContraPartida, iContaPreenchida)
        If lErro <> SUCESSO Then Error 11670
    End If

    If sContaResultado = sContaContraPartida Then Error 11650

    sContaResultadoNivel1 = String(STRING_CONTA, 0)

    'guarda o número do nivel 1 da conta resultado
    lErro = Mascara_RetornaContaNoNivel(1, sContaResultado, sContaResultadoNivel1)
    If lErro <> SUCESSO Then Error 9806
            
    If Len(Trim(ContaContraPartida.ClipText)) <> 0 Then
        sContaContraPartidaNivel1 = String(STRING_CONTA, 0)
    
        'guarda o número do nivel 1 da conta de contra partida
        lErro = Mascara_RetornaContaNoNivel(1, sContaContraPartida, sContaContraPartidaNivel1)
        If lErro <> SUCESSO Then Error 9807
    End If
            
    lErro = CF("PlanoConta_Le_Todas_Categorias", colPlanoConta)
    If lErro <> SUCESSO Then Error 9808

    lErro = CF("ContaCategoria_Le_Todos", colContaCategoria)
    If lErro <> SUCESSO Then Error 9809
    
    'Pesquisa as contas cuja categoria indique que participa da apuração
    For Each objPlanoConta In colPlanoConta
    
        For Each objContaCategoria In colContaCategoria
        
            If objPlanoConta.iCategoria = objContaCategoria.iCodigo Then
            
                If objContaCategoria.iApuracao = CONTACATEGORIA_APURACAO Then
                                    
                    'verifica se o grupo de contas a serem apuradas abrange a conta resultado
                    If objPlanoConta.sConta = sContaResultadoNivel1 Then Error 9804
                    
                    'verifica se o grupo de contas a serem apuradas abrange a conta de contra partida
                    If objPlanoConta.sConta = sContaContraPartidaNivel1 Then Error 9805
                    
                    colContasApuracao.Add objPlanoConta.sConta
                    
                    Exit For

                End If
                
            End If
            
        Next
        
    Next

    If colContasApuracao.Count = 0 Then Error 62420

    sHistorico = Historico.Text

    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 20343
    
    lErro = CF("Rotina_Apura_Periodos", sNomeArqParam, giFilialEmpresa, iExercicio, iPeriodoInicial, iPeriodoFinal, sContaResultado, sContaContraPartida, colContasApuracao, sHistorico, iZeraRD)
    If lErro <> SUCESSO Then Error 9800

    Call Limpa_Tela_ApPeriodo

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoApurar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 9798, 9800, 9808, 9809, 11670, 20343
        
        Case 9804
            Call Rotina_Erro(vbOKOnly, "ERRO_INTERSECAO_CONTARESULTADO_APURACAO", Err)
    
        Case 9805
            Call Rotina_Erro(vbOKOnly, "ERRO_INTERSECAO_CONTRAPARTIDA_APURACAO", Err)
    
        Case 9806
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, sContaResultado, 1)
    
        Case 9807
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, sContaContraPartida, 1)
    
        Case 11641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_SELECIONADO", Err)
            
        Case 11642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_NAO_SELECIONADO", Err)
        
        Case 11643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_FINAL_NAO_SELECIONADO", Err)
            
        Case 11644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_MAIOR", Err)
            
        Case 11645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACONTRAPARTIDA_VAZIA", Err)
        
        Case 11646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTARESULTADO_VAZIA", Err, iLinha)
                    
        Case 11650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_IGUAL_RESULTADO", Err)
            
        Case 62420
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_CONTAS_CATEGORIA_APURACAO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143126)
    
    End Select

    Exit Sub

End Sub

Private Sub ContaContraPartida_GotFocus()

    TvwContas.Visible = True
    LblTvws.Visible = True
    Historicos.Visible = False
    Hist.Visible = False
    
End Sub

Private Sub ContaResultado_GotFocus()

    TvwContas.Visible = True
    LblTvws.Visible = True
    Historicos.Visible = False
    Hist.Visible = False

End Sub

Private Sub ContaContraPartida_Validate(Cancel As Boolean)

Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long

On Error GoTo Erro_ContaContraPartida_Validate

    TvwContas.Tag = CONTA_CONTRA_PARTIDA

    If Len(Trim(ContaContraPartida.ClipText)) > 0 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContraPartida.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 11676

        'conta não cadastrada
        If lErro = 5700 Then Error 11677

    End If

    Exit Sub

Erro_ContaContraPartida_Validate:

    Cancel = True


    Select Case Err

        Case 11676

        Case 11677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143127)

    End Select

    Exit Sub

End Sub

Private Sub ContaResultado_Validate(Cancel As Boolean)

Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long

On Error GoTo Erro_ContaResultado_Validate

    TvwContas.Tag = CONTA_RESULTADO

    If Len(Trim(ContaResultado.ClipText)) > 0 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaResultado.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 9795

        'conta não cadastrada
        If lErro = 5700 Then Error 9796

    End If

    Exit Sub

Erro_ContaResultado_Validate:

    Cancel = True


    Select Case Err

        Case 9795

        Case 9796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143128)

    End Select

    Exit Sub

End Sub


Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Function Preenche_ComboExercicio(iPosInicial As Integer) As Long
'preenche Combo de Exercicios

Dim colExercicios As New Collection
Dim lErro As Long
Dim iConta As Integer
Dim objExercicio As ClassExercicio

On Error GoTo Erro_Preenche_ComboExercicio

    iPosInicial = 1

    'le todos os exercícios existentes no BD
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then Error 11678

    'preenche ComboBox com NomeExterno e ItemData com Exercicio
    For iConta = 1 To colExercicios.Count

        Set objExercicio = colExercicios.Item(iConta)
        If objExercicio.iStatus <> EXERCICIO_FECHADO Then
            Exercicio.AddItem objExercicio.sNomeExterno
            Exercicio.ItemData(Exercicio.NewIndex) = objExercicio.iExercicio
            If objExercicio.iExercicio = giExercicioAtual Then
                iPosInicial = Exercicio.NewIndex
            End If
        End If
    Next

    Preenche_ComboExercicio = SUCESSO

    Exit Function

Erro_Preenche_ComboExercicio:

    Preenche_ComboExercicio = Err

    Select Case Err

        Case 11678

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143129)

    End Select

    Exit Function

End Function

Private Function Carga_ListBox_Historico() As Long
'move os dados de historico padrão do banco de dados para a arvore colNodes.

Dim colHistPadrao As New Collection
Dim objHistPadrao As ClassHistPadrao
Dim sListBoxItem As String
Dim lErro As Long

On Error GoTo Erro_Carga_ListBox_Historico

    'Preenche a ListBox com Históricos Padrões existentes no BD
    lErro = CF("HistPadrao_Le_Todos", colHistPadrao)
    If lErro <> SUCESSO Then Error 11682
    
    For Each objHistPadrao In colHistPadrao
    
        'Espaços que faltam para completar tamanho STRING_CODIGO_HISTORICO
        sListBoxItem = Space(STRING_CODIGO_HISTORICO - Len(CStr(objHistPadrao.iHistPadrao)))
        
        'Concatena Codigo e Nome do HistPadrao
        sListBoxItem = sListBoxItem & CStr(objHistPadrao.iHistPadrao)
        sListBoxItem = sListBoxItem & SEPARADOR & objHistPadrao.sDescHistPadrao
    
        Historicos.AddItem sListBoxItem
        Historicos.ItemData(Historicos.NewIndex) = objHistPadrao.iHistPadrao
        
    Next
    
    Carga_ListBox_Historico = SUCESSO

    Exit Function

Erro_Carga_ListBox_Historico:

    Carga_ListBox_Historico = Err

    Select Case Err

        Case 11682

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143130)

    End Select

    Exit Function

End Function

Private Function Preenche_CombosPeriodos(iExercicio As Integer) As Long

Dim lErro As Long
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo
Dim iIndice As Integer

On Error GoTo Erro_Preenche_CombosPeriodos

    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 11683

    PeriodoInicial.Clear
    PeriodoFinal.Clear

    For Each objPeriodo In colPeriodos

        PeriodoInicial.AddItem objPeriodo.sNomeExterno
        PeriodoInicial.ItemData(PeriodoInicial.NewIndex) = objPeriodo.iPeriodo
        PeriodoFinal.AddItem objPeriodo.sNomeExterno
        PeriodoFinal.ItemData(PeriodoFinal.NewIndex) = objPeriodo.iPeriodo

    Next

    Preenche_CombosPeriodos = SUCESSO
    
    Exit Function

Erro_Preenche_CombosPeriodos:

    Preenche_CombosPeriodos = Err
    
    Select Case Err

        Case 11683

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143131)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_ApPeriodo()

    Exercicio.ListIndex = -1
    PeriodoInicial.ListIndex = -1
    PeriodoFinal.ListIndex = -1
    ContaContraPartida.PromptInclude = False
    ContaContraPartida.Text = ""
    ContaContraPartida.PromptInclude = True
    ContaResultado.PromptInclude = False
    ContaResultado.Text = ""
    ContaResultado.PromptInclude = True
    Historico.Text = ""
    PeriodoInicial.Clear
    PeriodoFinal.Clear

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

   Call Limpa_Tela_ApPeriodo

End Sub

Private Sub Historico_GotFocus()

    TvwContas.Visible = False
    LblTvws.Visible = False
    Historicos.Visible = True
    Hist.Visible = True

End Sub

Private Sub Historicos_DblClick()

Dim lPosicaoSeparador As Long
Dim lErro As Long

On Error GoTo Erro_Historicos_DblClick
    
    lPosicaoSeparador = InStr(Historicos.Text, SEPARADOR)
    Historico.Text = Mid(Historicos.Text, lPosicaoSeparador + 1)
 
    Exit Sub
    
Erro_Historicos_DblClick:

    Select Case Err
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143132)

    End Select
    
    Exit Sub
    
End Sub

Private Function Inicializa_Mascaras() As Long
'inicializa as mascaras de conta e centro de custo

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)

    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 11684

    ContaContraPartida.Mask = sMascaraConta
    ContaResultado.Mask = sMascaraConta

    Inicializa_Mascaras = SUCESSO

    Exit Function

Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err

    Select Case Err

        Case 11684

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143133)

    End Select

    Exit Function

End Function

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 40810
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 40810
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143134)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaMascarada As String
Dim cControl As Control
Dim iLinha As Integer

On Error GoTo Erro_TvwContas_NodeClick

    sCaracterInicial = left(Node.Key, 1)

    If sCaracterInicial <> "A" And (TvwContas.Tag = CONTA_CONTRA_PARTIDA Or TvwContas.Tag = CONTA_RESULTADO) Then Error 20301
    
    sConta = right(Node.Key, Len(Node.Key) - 1)

    sContaEnxuta = String(STRING_CONTA, 0)

    'volta mascarado apenas os caracteres preenchidos
    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 20302

    If TvwContas.Tag = CONTA_RESULTADO Then

        ContaResultado.PromptInclude = False
        ContaResultado.Text = sContaEnxuta
        ContaResultado.PromptInclude = True

    ElseIf TvwContas.Tag = CONTA_CONTRA_PARTIDA Then

        ContaContraPartida.PromptInclude = False
        ContaContraPartida.Text = sContaEnxuta
        ContaContraPartida.PromptInclude = True

    End If

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 20301

        Case 20302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143135)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_APURACAO_PERIODO
    Set Form_Load_Ocx = Me
    Caption = "Apuração de Periodo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ApuraPeriodo"
    
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




Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Hist_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Hist, Source, X, Y)
End Sub

Private Sub Hist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Hist, Button, Shift, X, Y)
End Sub

Private Sub LblTvws_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTvws, Source, X, Y)
End Sub

Private Sub LblTvws_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTvws, Button, Shift, X, Y)
End Sub

Private Sub ZeraRD_Click()
    If ZeraRD.Value = vbChecked Then
        ContaContraPartida.Enabled = False
        ContaContraPartida.PromptInclude = False
        ContaContraPartida.Text = ""
        ContaContraPartida.PromptInclude = True
        Label1.Enabled = False
    Else
        ContaContraPartida.Enabled = False
        Label1.Enabled = False
    End If
End Sub
