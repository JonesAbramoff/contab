VERSION 5.00
Begin VB.UserControl RelOpLoteOcx 
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   LockControls    =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   6990
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4680
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpLoteOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpLoteOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpLoteOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpLoteOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpLoteOcx.ctx":0994
      Left            =   1515
      List            =   "RelOpLoteOcx.ctx":0996
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   975
      Width           =   1590
   End
   Begin VB.ComboBox ComboPeriodoInic 
      Height          =   315
      ItemData        =   "RelOpLoteOcx.ctx":0998
      Left            =   1515
      List            =   "RelOpLoteOcx.ctx":099A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1755
      Width           =   1935
   End
   Begin VB.TextBox LoteFinal 
      Height          =   315
      Left            =   4950
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox LoteInicial 
      Height          =   315
      Left            =   1515
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpLoteOcx.ctx":099C
      Left            =   1515
      List            =   "RelOpLoteOcx.ctx":099E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2475
   End
   Begin VB.ComboBox ComboPeriodoFim 
      Height          =   315
      ItemData        =   "RelOpLoteOcx.ctx":09A0
      Left            =   4935
      List            =   "RelOpLoteOcx.ctx":09A2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1755
      Width           =   1935
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4815
      Picture         =   "RelOpLoteOcx.ctx":09A4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   855
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   17
      Top             =   1815
      Width           =   1320
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   585
      TabIndex        =   16
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Lote Final:"
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
      Left            =   3945
      TabIndex        =   15
      Top             =   2565
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Lote Inicial:"
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
      Left            =   420
      TabIndex        =   14
      Top             =   2565
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Opção:"
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
      Left            =   795
      TabIndex        =   13
      Top             =   300
      Width           =   630
   End
   Begin VB.Label Label4 
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
      Left            =   3660
      TabIndex        =   12
      Top             =   1815
      Width           =   1215
   End
End
Attribute VB_Name = "RelOpLoteOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giCarregando As Integer

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção
'recebe os lotes inicial e final no formato do BD

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If LoteInicial.Text <> "" Then sExpressao = "Lote >= " & Forprint_ConvInt(CInt(LoteInicial.Text))

    If LoteFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Lote <= " & Forprint_ConvInt(CInt(LoteFinal.Text))

    End If

    Select Case giFilialEmpresa
        Case EMPRESA_TODA
            If giContabGerencial <> 0 Then
                If sExpressao <> "" Then sExpressao = sExpressao & " E "
                sExpressao = sExpressao & "FilialEmpresa < " & Forprint_ConvInt(Abs(giFilialAuxiliar))
            End If
        
        Case Abs(giFilialAuxiliar)
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "FilialEmpresa > " & Forprint_ConvInt(Abs(giFilialAuxiliar))
        
        Case Else
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
    End Select
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169861)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim iExercicio As Integer, iPer_I As Integer, iPer_F As Integer
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 13506

    'pega Lote Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NLOTEINIC", sParam)
    If lErro <> SUCESSO Then Error 13507

    LoteInicial.Text = sParam

    'pega Lote Final e exibe
    lErro = objRelOpcoes.ObterParametro("NLOTEFIM", sParam)
    If lErro <> SUCESSO Then Error 13508

    LoteFinal.Text = sParam

    'período inicial
    lErro = objRelOpcoes.ObterParametro("NPERINIC", sParam)
    If lErro <> SUCESSO Then Error 13509

    iPer_I = CInt(sParam)

    'período final
    lErro = objRelOpcoes.ObterParametro("NPERFIM", sParam)
    If lErro <> SUCESSO Then Error 13510

    iPer_F = CInt(sParam)

    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 13511

    iExercicio = CInt(sParam)

    lErro = MostraExercicioPeriodos(iExercicio, iPer_I, iPer_F)
    If lErro <> SUCESSO Then Error 13512

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 13506

        Case 13507, 13508, 13509, 13510, 13511

        Case 13512

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169862)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sLote_I As String, sLote_F As String

On Error GoTo Erro_PreencherRelOp

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 13514

    'período inicial não pode ser vazio
    If ComboPeriodoInic.Text = "" Then Error 13515

    'período final não pode ser vazio
    If ComboPeriodoFim.Text = "" Then Error 13516

    'período inicial não pode ser maior que o período final
    If ComboPeriodoInic.ItemData(ComboPeriodoInic.ListIndex) > ComboPeriodoFim.ItemData(ComboPeriodoFim.ListIndex) Then Error 13542

    'lote inicial não pode ser maior que o lote final
    If LoteInicial.Text <> "" And LoteFinal.Text <> "" Then
        If CInt(LoteInicial.Text) > CInt(LoteFinal.Text) Then Error 13517
    End If

    'grava os parâmetros no arquivo C
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 13518

    lErro = objRelOpcoes.IncluirParametro("NLOTEINIC", LoteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13519

    lErro = objRelOpcoes.IncluirParametro("NLOTEFIM", LoteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13520

    lErro = objRelOpcoes.IncluirParametro("NPERINIC", CStr(ComboPeriodoInic.ItemData(ComboPeriodoInic.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13521

    lErro = objRelOpcoes.IncluirParametro("NPERFIM", CStr(ComboPeriodoFim.ItemData(ComboPeriodoFim.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13522

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13523
        
    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13493

    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", ComboPeriodoInic.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13541

    lErro = objRelOpcoes.IncluirParametro("TTITAUX3", ComboPeriodoFim.Text)
    If lErro <> AD_BOOL_TRUE Then Error 20599

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 13524

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 13514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 13515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_VAZIO", Err)
            ComboPeriodoInic.SetFocus

        Case 13516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_FINAL_VAZIO", Err)
            ComboPeriodoFim.SetFocus

        Case 13517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INICIAL_MAIOR", Err)

        Case 13542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_MAIOR", Err)

        Case 13493, 13518

        Case 13519, 13520, 13521, 13522, 13523

        Case 13524, 13541, 20599

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169863)

    End Select

    Exit Function

End Function

Function MostraExercicioPeriodos(iExercicio As Integer, iPer_I As Integer, iPer_F As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama PreencheComboPeriodo

Dim iConta As Integer, lErro As Long

On Error GoTo Erro_MostraExercicioPeriodos

    giCarregando = OK

    For iConta = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iConta) = iExercicio Then
            ComboExercicio.ListIndex = iConta
            Exit For
        End If
    Next

    lErro = PreencheComboPeriodos(iExercicio, iPer_I, iPer_F)
    If lErro <> SUCESSO Then Error 13526

    MostraExercicioPeriodos = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodos:

    MostraExercicioPeriodos = Err

    Select Case Err

        Case 13526

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169864)

    End Select

    Exit Function

End Function

Function PreencheComboPeriodos(iExercicio As Integer, iPer_I As Integer, iPer_F As Integer) As Long
'lê os períodos do exercício 'iExercicio' preenchendo o combo de período
'seleciona o período 'iPeriodo'

Dim lErro As Long, iConta As Integer
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheComboPeriodos

    ComboPeriodoInic.Clear
    ComboPeriodoFim.Clear

    'inicializar os periodos do exercicio selecionado no combo de exercícios
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 13527

    For iConta = 1 To colPeriodos.Count
        
        Set objPeriodo = colPeriodos.Item(iConta)
        
        ComboPeriodoInic.AddItem objPeriodo.sNomeExterno
        ComboPeriodoInic.ItemData(ComboPeriodoInic.NewIndex) = objPeriodo.iPeriodo
        
        ComboPeriodoFim.AddItem objPeriodo.sNomeExterno
        ComboPeriodoFim.ItemData(ComboPeriodoFim.NewIndex) = objPeriodo.iPeriodo
    
    Next

    'mostra o período inicial
    For iConta = 0 To ComboPeriodoInic.ListCount - 1
        If ComboPeriodoInic.ItemData(iConta) = iPer_I Then
            ComboPeriodoInic.ListIndex = iConta
            Exit For
        End If
    Next

    'mostra o período final
    For iConta = 0 To ComboPeriodoFim.ListCount - 1
        If ComboPeriodoFim.ItemData(iConta) = iPer_F Then
            ComboPeriodoFim.ListIndex = iConta
            Exit For
        End If
    Next
    
    PreencheComboPeriodos = SUCESSO

    Exit Function

Erro_PreencheComboPeriodos:

    PreencheComboPeriodos = Err

    Select Case Err

        Case 13527

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169865)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29561
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 13536

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 13536
        
        Case 29561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169866)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 13528

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 13529

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47074
    
    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 13528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 13529, 47074

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169867)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13530

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 13530

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169868)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 13513

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13531

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 13532

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47075
    
    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 13513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 13531

        Case 13532, 47075
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169869)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47073
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47073
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169870)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click

    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    If giCarregando = CANCELA Then

        lErro = PreencheComboPeriodos(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1, 1)
        If lErro <> SUCESSO Then Error 13533
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 13533

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169871)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long, iConta As Integer
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection

On Error GoTo Erro_Form_Load

    giCarregando = CANCELA
    
    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 13537
    
    For iConta = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iConta)
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next

    'verifica se o nome da opção passada está no ComboBox
    For iConta = 0 To ComboOpcoes.ListCount - 1

        If ComboOpcoes.List(iConta) = gobjRelOpcoes.sNome Then

            ComboOpcoes.Text = ComboOpcoes.List(iConta)
            PreencherParametrosNaTela (gobjRelOpcoes)

            Exit For

        End If

    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 13537

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169872)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub LoteFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteFinal_Validate

    'lote final deve estar entre 1 e 9999
    If LoteFinal.Text <> "" Then
        lErro = Valor_Critica(LoteFinal.Text)
        If lErro <> SUCESSO Then Error 54874

        If (CInt(LoteFinal.Text) < 1) Or (CInt(LoteFinal.Text) > 9999) Then Error 13539
    End If

    Exit Sub

Erro_LoteFinal_Validate:

    Cancel = True


    Select Case Err
        
        Case 54874

        Case 13539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FORA_FAIXA", Err)
            LoteFinal.Text = ""

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169873)

    End Select

    Exit Sub

End Sub

Private Sub LoteInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteInicial_Validate

    'lote inicial deve estar entre 1 e 9999
    If LoteInicial.Text <> "" Then
        lErro = Valor_Critica(LoteInicial.Text)
        If lErro <> SUCESSO Then Error 54873

        If (CInt(LoteInicial.Text) < 1) Or (CInt(LoteInicial.Text) > 9999) Then Error 13540
    End If

    Exit Sub

Erro_LoteInicial_Validate:

    Cancel = True


    Select Case Err
        
        Case 54873

        Case 13540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FORA_FAIXA", Err)
            LoteInicial.Text = ""

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169874)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_LOTE
    Set Form_Load_Ocx = Me
    Caption = "Listagem de Lotes Contabilizados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpLote"
    
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

Public Sub Unload(objme As Object)
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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
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

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

