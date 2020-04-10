VERSION 5.00
Begin VB.UserControl RelOpLotePendOcx 
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   6975
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
         Picture         =   "RelOpLotePendOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpLotePendOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpLotePendOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpLotePendOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
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
      Picture         =   "RelOpLotePendOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   870
      Width           =   1815
   End
   Begin VB.ComboBox ComboPeriodoFim 
      Height          =   315
      Left            =   4935
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1755
      Width           =   1935
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpLotePendOcx.ctx":0A96
      Left            =   1515
      List            =   "RelOpLotePendOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2475
   End
   Begin VB.TextBox LoteInicial 
      Height          =   315
      Left            =   1515
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox LoteFinal 
      Height          =   315
      Left            =   4950
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ComboBox ComboPeriodoInic 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1770
      Width           =   1935
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpLotePendOcx.ctx":0A9A
      Left            =   1515
      List            =   "RelOpLotePendOcx.ctx":0A9C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   975
      Width           =   1590
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
      TabIndex        =   17
      Top             =   1815
      Width           =   1215
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
      TabIndex        =   16
      Top             =   300
      Width           =   630
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
      TabIndex        =   15
      Top             =   2565
      Width           =   1020
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
      TabIndex        =   14
      Top             =   2565
      Width           =   915
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
      TabIndex        =   13
      Top             =   1020
      Width           =   855
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
      TabIndex        =   12
      Top             =   1815
      Width           =   1320
   End
End
Attribute VB_Name = "RelOpLotePendOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169875)

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
    If lErro <> SUCESSO Then Error 40949

    'pega Lote Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NLOTEINIC", sParam)
    If lErro <> SUCESSO Then Error 40950

    LoteInicial.Text = sParam

    'pega Lote Final e exibe
    lErro = objRelOpcoes.ObterParametro("NLOTEFIM", sParam)
    If lErro <> SUCESSO Then Error 40951

    LoteFinal.Text = sParam

    'período inicial
    lErro = objRelOpcoes.ObterParametro("NPERINIC", sParam)
    If lErro <> SUCESSO Then Error 40952

    iPer_I = CInt(sParam)

    'período final
    lErro = objRelOpcoes.ObterParametro("NPERFIM", sParam)
    If lErro <> SUCESSO Then Error 40953

    iPer_F = CInt(sParam)

    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 40954

    iExercicio = CInt(sParam)

    lErro = MostraExercicioPeriodos(iExercicio, iPer_I, iPer_F)
    If lErro <> SUCESSO Then Error 40955

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 40949, 40950, 40951, 40952, 40953, 40954, 40955

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169876)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sLote_I As String, sLote_F As String

On Error GoTo Erro_PreencherRelOp

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 40956

    'período inicial não pode ser vazio
    If ComboPeriodoInic.Text = "" Then Error 40957

    'período final não pode ser vazio
    If ComboPeriodoFim.Text = "" Then Error 40958

    'período inicial não pode ser maior que o período final
    If ComboPeriodoInic.ItemData(ComboPeriodoInic.ListIndex) > ComboPeriodoFim.ItemData(ComboPeriodoFim.ListIndex) Then Error 40959

    'lote inicial não pode ser maior que o lote final
    If LoteInicial.Text <> "" And LoteFinal.Text <> "" Then
        If CInt(LoteInicial.Text) > CInt(LoteFinal.Text) Then Error 40960
    End If

    'grava os parâmetros no arquivo C
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 40961

    lErro = objRelOpcoes.IncluirParametro("NLOTEINIC", LoteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 40962

    lErro = objRelOpcoes.IncluirParametro("NLOTEFIM", LoteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 40963

    lErro = objRelOpcoes.IncluirParametro("NPERINIC", CStr(ComboPeriodoInic.ItemData(ComboPeriodoInic.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 40964

    lErro = objRelOpcoes.IncluirParametro("NPERFIM", CStr(ComboPeriodoFim.ItemData(ComboPeriodoFim.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 40965

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 40966
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47070

    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", ComboPeriodoInic.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47071

    lErro = objRelOpcoes.IncluirParametro("TTITAUX3", ComboPeriodoFim.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47072


    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 40967

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 40956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 40957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_VAZIO", Err)
            ComboPeriodoInic.SetFocus

        Case 40958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_FINAL_VAZIO", Err)
            ComboPeriodoFim.SetFocus
        
        Case 40959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_MAIOR", Err)

        Case 40960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INICIAL_MAIOR", Err)

        Case 40961, 40962, 40963, 40964, 40965, 40966, 40967, 47070, 47071, 47072
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 169877)

    End Select

    Exit Function

End Function

Function MostraExercicioPeriodos(iExercicio As Integer, iPer_I As Integer, iPer_F As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama PreencheComboPeriodo

Dim iIndice As Integer, lErro As Long

On Error GoTo Erro_MostraExercicioPeriodos

    giCarregando = OK

    For iIndice = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iIndice) = iExercicio Then
            ComboExercicio.ListIndex = iIndice
            Exit For
        End If
    Next

    lErro = PreencheComboPeriodos(iExercicio, iPer_I, iPer_F)
    If lErro <> SUCESSO Then Error 40969

    MostraExercicioPeriodos = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodos:

    MostraExercicioPeriodos = Err

    Select Case Err

        Case 40969

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169878)

    End Select

    Exit Function

End Function

Function PreencheComboPeriodos(iExercicio As Integer, iPer_I As Integer, iPer_F As Integer) As Long
'lê os períodos do exercício 'iExercicio' preenchendo o combo de período
'seleciona o período 'iPeriodo'

Dim lErro As Long
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheComboPeriodos

    ComboPeriodoInic.Clear
    ComboPeriodoFim.Clear

    'inicializar os periodos do exercicio selecionado no combo de exercícios
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 40970

    For Each objPeriodo In colPeriodos

        ComboPeriodoInic.AddItem objPeriodo.sNomeExterno
        ComboPeriodoInic.ItemData(ComboPeriodoInic.NewIndex) = objPeriodo.iPeriodo

        If ComboPeriodoInic.ItemData(ComboPeriodoInic.NewIndex) = iPer_I Then
            ComboPeriodoInic.ListIndex = ComboPeriodoInic.NewIndex
        End If

        ComboPeriodoFim.AddItem objPeriodo.sNomeExterno
        ComboPeriodoFim.ItemData(ComboPeriodoFim.NewIndex) = objPeriodo.iPeriodo

        If ComboPeriodoFim.ItemData(ComboPeriodoFim.NewIndex) = iPer_F Then
            ComboPeriodoFim.ListIndex = ComboPeriodoFim.NewIndex
        End If

    Next

    PreencheComboPeriodos = SUCESSO

    Exit Function

Erro_PreencheComboPeriodos:

    PreencheComboPeriodos = Err

    Select Case Err

        Case 40970

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169879)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29560
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 40980

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 40980
        
        Case 29560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169880)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 40971

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 40972

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47077

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 40971
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 40972, 47077

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169881)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 40973

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 40973

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169882)

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
    If ComboOpcoes.Text = "" Then Error 40974

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 40975

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 40976

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47078
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 40974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 40975, 40976, 47078, 47078

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169883)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47080
    
    ComboOpcoes.Text = ""
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47080
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169884)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click

    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    If giCarregando = CANCELA Then

        lErro = PreencheComboPeriodos(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1, 1)
        If lErro <> SUCESSO Then Error 40977
    
    End If
    
    giCarregando = CANCELA

    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 40977

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169885)

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
    lErro = CF("Exercicios_Abertos_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 40981

    For Each objExercicio In colExerciciosAbertos
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next

    lErro = MostraExercicioPeriodos(giExercicioAtual, giPeriodoAtual, giPeriodoAtual)
    If lErro <> SUCESSO Then Error 40982

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

        Case 40981, 40982

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169886)

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
        If lErro <> SUCESSO Then Error 54876

        If (CInt(LoteFinal.Text) < 1) Or (CInt(LoteFinal.Text) > 9999) Then Error 40983
    End If

    Exit Sub

Erro_LoteFinal_Validate:

    Cancel = True


    Select Case Err
        
        Case 54876

        Case 40983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FORA_FAIXA", Err)
            LoteFinal.Text = ""

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169887)

    End Select

    Exit Sub

End Sub

Private Sub LoteInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteInicial_Validate

    'lote inicial deve estar entre 1 e 9999
    If LoteInicial.Text <> "" Then
        lErro = Valor_Critica(LoteInicial.Text)
        If lErro <> SUCESSO Then Error 54875

        If (CInt(LoteInicial.Text) < 1) Or (CInt(LoteInicial.Text) > 9999) Then Error 40984
    End If

    Exit Sub

Erro_LoteInicial_Validate:

    Cancel = True


    Select Case Err
        
        Case 54875

        Case 40984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_FORA_FAIXA", Err)
            LoteInicial.Text = ""

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169888)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_LOTE_PENDENTE
    Set Form_Load_Ocx = Me
    Caption = "Listagem de Lotes Pendentes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpLotePend"
    
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




Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

