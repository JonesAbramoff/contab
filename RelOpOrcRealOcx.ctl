VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpOrcRealOcx 
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   KeyPreview      =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   8310
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpOrcRealOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOrcRealOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpOrcRealOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpOrcRealOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpOrcRealOcx.ctx":0994
      Left            =   1035
      List            =   "RelOpOrcRealOcx.ctx":0996
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   915
      Width           =   1380
   End
   Begin VB.ComboBox ComboPeriodo 
      Height          =   315
      ItemData        =   "RelOpOrcRealOcx.ctx":0998
      Left            =   3555
      List            =   "RelOpOrcRealOcx.ctx":099A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   915
      Width           =   1695
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOrcRealOcx.ctx":099C
      Left            =   1035
      List            =   "RelOpOrcRealOcx.ctx":099E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contas"
      Height          =   1665
      Left            =   165
      TabIndex        =   12
      Top             =   1455
      Width           =   7980
      Begin MSMask.MaskEdBox ContaInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   435
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaFinal 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   1050
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescCtaFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2740
         TabIndex        =   16
         Top             =   1050
         Width           =   5000
      End
      Begin VB.Label DescCtaInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2740
         TabIndex        =   15
         Top             =   435
         Width           =   5000
      End
      Begin VB.Label LabelContaDe 
         Caption         =   "Inicial:"
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
         Height          =   255
         Left            =   75
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   435
         Width           =   615
      End
      Begin VB.Label LabelContaAte 
         Caption         =   "Final:"
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
         Height          =   255
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   1065
         Width           =   615
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
      Left            =   4155
      Picture         =   "RelOpOrcRealOcx.ctx":09A0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   2760
      Left            =   5730
      TabIndex        =   5
      Top             =   1110
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   4868
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2760
      TabIndex        =   20
      Top             =   975
      Width           =   750
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
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   855
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
      Left            =   330
      TabIndex        =   18
      Top             =   285
      Width           =   630
   End
   Begin VB.Label LabelContas 
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
      Left            =   5745
      TabIndex        =   17
      Top             =   870
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "RelOpOrcRealOcx"
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
Dim giFocoInicial As Integer
Dim giCarregando As Integer

Private WithEvents objEventoContaDe As AdmEvento
Attribute objEventoContaDe.VB_VarHelpID = -1
Private WithEvents objEventoContaAte As AdmEvento
Attribute objEventoContaAte.VB_VarHelpID = -1

Private Sub BotaoExcluir_Click()
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 40906

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 40907

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47082
        
        ComboOpcoes.Text = ""
        DescCtaFim.Caption = ""
        DescCtaInic.Caption = ""
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 40906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 40907, 47082

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170425)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 40908

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 40908

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170426)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()
Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 40909

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 40910

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 40911

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47083
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 40909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 40910, 40911, 47083

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170427)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47081
    
    ComboOpcoes.Text = ""
    DescCtaFim.Caption = ""
    DescCtaInic.Caption = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47081
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170428)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click

    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    If giCarregando = CANCELA Then
    
        lErro = PreencheComboPeriodo(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1)
        If lErro <> SUCESSO Then Error 40912
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 40912

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170429)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub ContaFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ContaFinal_Validate

    giFocoInicial = 0

    lErro = CF("Conta_Perde_Foco", ContaFinal, DescCtaFim)
    If lErro <> SUCESSO Then Error 40941

    Exit Sub

Erro_ContaFinal_Validate:

    Cancel = True


    Select Case Err

        Case 40941

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170430)

    End Select

    Exit Sub

End Sub

Private Sub ContaInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ContaInicial_Validate

    giFocoInicial = 1

    lErro = CF("Conta_Perde_Foco", ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then Error 40942

    Exit Sub

Erro_ContaInicial_Validate:

    Cancel = True


    Select Case Err

        Case 40942

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170431)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long, iConta As Integer
Dim objExercicio As ClassExercicio
Dim colExercicios As New Collection

On Error GoTo Erro_Form_Load

    giCarregando = CANCELA
    
    'inicializa a mascara de conta inicial
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaInicial)
    If lErro <> SUCESSO Then Error 40943
    
    'inicializa a mascara de conta final
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaFinal)
    If lErro <> SUCESSO Then Error 40944

    'ler os exercicios
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then Error 40946

    'preenche a Combo de exercicio
    For Each objExercicio In colExercicios
        
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
        
        If ComboExercicio.ItemData(ComboExercicio.NewIndex) = giExercicioAtual Then
            ComboExercicio.ListIndex = ComboExercicio.NewIndex
        End If
    Next
'
'     'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 40947

    Set objEventoContaDe = New AdmEvento
    Set objEventoContaAte = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 40943, 40944, 40946, 40947

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170432)

    End Select

    Unload Me

    Exit Sub
    
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then

        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 40915

    End If

    Exit Sub

Erro_TvwContas_Expand:

    Select Case Err

        Case 40915

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170433)

    End Select

    Exit Sub

End Sub
Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_TvwContas_NodeClick

    sConta = Right(Node.Key, Len(Node.Key) - 1)

    lErro = Traz_Conta_Tela(sConta)
    If lErro <> SUCESSO Then Error 40916

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 40916

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170434)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCta_I As String, sCta_F As String

On Error GoTo Erro_PreencherRelOp

    sCta_I = String(STRING_CONTA, 0)
    sCta_F = String(STRING_CONTA, 0)
    
    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 40917

    'período inicial não pode ser vazio
    If ComboPeriodo.Text = "" Then Error 40918
 
    lErro = Formata_E_Critica_Contas(sCta_I, sCta_F)
    If lErro <> SUCESSO Then Error 40919

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 40920

    lErro = objRelOpcoes.IncluirParametro("TCTAINIC", sCta_I)
    If lErro <> AD_BOOL_TRUE Then Error 40921
    
     lErro = objRelOpcoes.IncluirParametro("TCTAFIM", sCta_F)
    If lErro <> AD_BOOL_TRUE Then Error 40922
    
    lErro = objRelOpcoes.IncluirParametro("NPERIODO", CStr(ComboPeriodo.ItemData(ComboPeriodo.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 40923
    
    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 40924
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 45091

    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", ComboPeriodo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 45092

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCta_I, sCta_F)
    If lErro <> SUCESSO Then Error 40925

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 40917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 40918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)
            ComboPeriodo.SetFocus
                
        Case 40919, 40920, 40921, 40922, 40922, 40923, 40924, 40925, 45091, 45092

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170435)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCta_I As String, sCta_F As String) As Long
'monta a expressão de seleção
'recebe as contas inicial e final no formato do BD

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sCta_I <> "" Then sExpressao = "Conta >= " & Forprint_ConvTexto(sCta_I)

    If sCta_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Conta <= " & Forprint_ConvTexto(sCta_F)

    End If
    
    If sExpressao <> "" Then
    
        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170436)

    End Select

    Exit Function

End Function

Function Formata_E_Critica_Contas(sCta_I As String, sCta_F As String) As Long
'Formata as contas retornando em sCta_I e sCta_F
'Verifica se a conta inicial é maior que a conta final
'função idêntica em RelOpBalVerif e RelOpRazao

Dim iCtaPreenchida_I As Integer, iCtaPreenchida_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Contas

    'formata a Conta Inicial
    lErro = CF("Conta_Formata", ContaInicial.Text, sCta_I, iCtaPreenchida_I)
    If lErro <> SUCESSO Then Error 40926
    If iCtaPreenchida_I <> CONTA_PREENCHIDA Then sCta_I = ""
    
    'formata a Conta Final
    lErro = CF("Conta_Formata", ContaFinal.Text, sCta_F, iCtaPreenchida_F)
    If lErro <> SUCESSO Then Error 40927
    If iCtaPreenchida_F <> CONTA_PREENCHIDA Then sCta_F = ""

    'se ambas as contas estão preenchidas, a conta inicial não pode ser maior que a final
    If iCtaPreenchida_I = CONTA_PREENCHIDA And iCtaPreenchida_F = CONTA_PREENCHIDA Then

        If sCta_I > sCta_F Then Error 40928

    End If

    Formata_E_Critica_Contas = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Contas:

    Formata_E_Critica_Contas = Err

    Select Case Err

        Case 40926
            ContaInicial.SetFocus

        Case 40927
            ContaFinal.SetFocus

        Case 40928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170437)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 40929

    'pega Conta Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAINIC", sParam)
    If lErro <> SUCESSO Then Error 40930

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then Error 40931

    'pega Conta Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAFIM", sParam)
    If lErro <> SUCESSO Then Error 40932

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaFinal, DescCtaFim)
    If lErro <> SUCESSO Then Error 40933

    'período inicial
    lErro = objRelOpcoes.ObterParametro("NPERIODO", sParam)
    If lErro <> SUCESSO Then Error 40934

    iPeriodo = CInt(sParam)
        
    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 40935

    iExercicio = CInt(sParam)
    
    'Traz a tela o periodo e o Exercicio da opcao desejada
    lErro = MostraExercicioPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 40936
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 40929, 40930, 40931, 40932, 40933, 40934, 40935, 40936

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170438)

    End Select

    Exit Function

End Function

Function MostraExercicioPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama PreencheComboPeriodo

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_MostraExercicioPeriodo

    giCarregando = OK

    For iIndice = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iIndice) = iExercicio Then
            ComboExercicio.ListIndex = iIndice
            Exit For
        End If
    Next

    lErro = PreencheComboPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 40937

    MostraExercicioPeriodo = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodo:

    MostraExercicioPeriodo = Err

    Select Case Err

        Case 40937

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170439)

    End Select

    Exit Function

End Function

Function PreencheComboPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'lê os períodos do exercício 'iExercicio' preenchendo o combo de período
'seleciona o período 'iPeriodo'

Dim lErro As Long
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheComboPeriodo

    ComboPeriodo.Clear
    
    'inicializar os periodos do exercicio selecionado no combo de exercícios
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 40938
    
    'adiciona os periodos na combo e posiciona com o item passado nos parametros iPer_I e iPer_F
    For Each objPeriodo In colPeriodos
    
        ComboPeriodo.AddItem objPeriodo.sNomeExterno
        ComboPeriodo.ItemData(ComboPeriodo.NewIndex) = objPeriodo.iPeriodo
        
        If ComboPeriodo.ItemData(ComboPeriodo.NewIndex) = iPeriodo Then
            ComboPeriodo.ListIndex = ComboPeriodo.NewIndex
        End If
                
    Next
  
    PreencheComboPeriodo = SUCESSO

    Exit Function

Erro_PreencheComboPeriodo:

    PreencheComboPeriodo = Err

    Select Case Err

        Case 40938

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170440)

    End Select

    Exit Function

End Function

Function Traz_Conta_Tela(sConta As String) As Long
'verifica e preenche a conta inicial e final com sua descriçao de acordo com o último foco
'sConta deve estar no formato do BD

Dim lErro As Long

On Error GoTo Erro_Traz_Conta_Tela

    If giFocoInicial Then

        lErro = CF("Traz_Conta_MaskEd", sConta, ContaInicial, DescCtaInic)
        If lErro <> SUCESSO Then Error 40939

    Else

        lErro = CF("Traz_Conta_MaskEd", sConta, ContaFinal, DescCtaFim)
        If lErro <> SUCESSO Then Error 40940

    End If

    Traz_Conta_Tela = SUCESSO

    Exit Function

Erro_Traz_Conta_Tela:

    Traz_Conta_Tela = Err

    Select Case Err

        Case 40939, 40940

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170441)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoContaDe = Nothing
    Set objEventoContaAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 24979
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 40945

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 40945
        
        Case 24979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170442)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ORC_REAL
    Set Form_Load_Ocx = Me
    Caption = "Orçados vs Reais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpOrcReal"
    
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


Private Sub DescCtaFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCtaFim, Source, X, Y)
End Sub

Private Sub DescCtaFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCtaFim, Button, Shift, X, Y)
End Sub

Private Sub DescCtaInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCtaInic, Source, X, Y)
End Sub

Private Sub DescCtaInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCtaInic, Button, Shift, X, Y)
End Sub

'Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label4, Source, X, Y)
'End Sub
'
'Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label5, Source, X, Y)
'End Sub
'
'Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
'End Sub

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

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Public Sub LabelContaDe_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim sContaOrigem As String
Dim iContaPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_LabelContaDe_Click

    If Len(Trim(ContaInicial.ClipText)) > 0 Then
    
        lErro = CF("Conta_Formata", ContaInicial.Text, sContaOrigem, iContaPreenchida)
        If lErro <> SUCESSO Then gError 197943

        If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sContaOrigem
    Else
        objPlanoConta.sConta = ""
    End If
           
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoContaDe)

    Exit Sub
    
Erro_LabelContaDe_Click:

    Select Case gErr
        
        Case 197943
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197945)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoContaDe_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaDe_evSelecao
    
    Set objPlanoConta = obj1
    
    sConta = objPlanoConta.sConta
    
    sContaEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then gError 197939

    ContaInicial.PromptInclude = False
    ContaInicial.Text = sContaEnxuta
    ContaInicial.PromptInclude = True
    Call ContaInicial_Validate(bSGECancelDummy)

    Me.Show
    
    Exit Sub
    
Erro_objEventoContaDe_evSelecao:

    Select Case gErr

        Case 197939
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197942)
        
    End Select

    Exit Sub

End Sub

Public Sub LabelContaAte_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim sContaOrigem As String
Dim iContaPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_LabelContaAte_Click

    If Len(Trim(ContaFinal.ClipText)) > 0 Then
    
        lErro = CF("Conta_Formata", ContaFinal.Text, sContaOrigem, iContaPreenchida)
        If lErro <> SUCESSO Then gError 197943

        If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sContaOrigem
    Else
        objPlanoConta.sConta = ""
    End If
           
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoContaAte)

    Exit Sub
    
Erro_LabelContaAte_Click:

    Select Case gErr
        
        Case 197943
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197945)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoContaAte_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaAte_evSelecao
    
    Set objPlanoConta = obj1
    
    sConta = objPlanoConta.sConta
    
    sContaEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then gError 197939

    ContaFinal.PromptInclude = False
    ContaFinal.Text = sContaEnxuta
    ContaFinal.PromptInclude = True
    Call ContaFinal_Validate(bSGECancelDummy)

    Me.Show
    
    Exit Sub
    
Erro_objEventoContaAte_evSelecao:

    Select Case gErr

        Case 197939
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197942)
        
    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ContaInicial Then Call LabelContaDe_Click
        If Me.ActiveControl Is ContaFinal Then Call LabelContaAte_Click
    
    End If
    
End Sub


