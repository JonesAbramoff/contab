VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpOrcRealCclOcx 
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   KeyPreview      =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   8325
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
         Picture         =   "RelOpOrcRealCclOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOrcRealCclOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpOrcRealCclOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpOrcRealCclOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
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
      Left            =   4155
      Picture         =   "RelOpOrcRealCclOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Centros de Custo"
      Height          =   1665
      Left            =   165
      TabIndex        =   12
      Top             =   1470
      Width           =   7980
      Begin MSMask.MaskEdBox CclInicial 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   465
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CclFinal 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   1065
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label LabelCclAte 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   1110
         Width           =   480
      End
      Begin VB.Label LabelCclDe 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   510
         Width           =   585
      End
      Begin VB.Label DescCclInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2730
         TabIndex        =   14
         Top             =   450
         Width           =   4995
      End
      Begin VB.Label DescCclFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2730
         TabIndex        =   13
         Top             =   1050
         Width           =   4995
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOrcRealCclOcx.ctx":0A96
      Left            =   1035
      List            =   "RelOpOrcRealCclOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox ComboPeriodo 
      Height          =   315
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   915
      Width           =   1695
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpOrcRealCclOcx.ctx":0A9A
      Left            =   1035
      List            =   "RelOpOrcRealCclOcx.ctx":0AA4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   1380
   End
   Begin MSComctlLib.TreeView TvwCcls 
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Centros de Custo"
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
      TabIndex        =   20
      Top             =   870
      Visible         =   0   'False
      Width           =   1470
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
      TabIndex        =   19
      Top             =   285
      Width           =   630
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
      TabIndex        =   18
      Top             =   945
      Width           =   855
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
      Left            =   2850
      TabIndex        =   17
      Top             =   960
      Width           =   750
   End
End
Attribute VB_Name = "RelOpOrcRealCclOcx"
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
Dim giFocoInicial As Integer

Private WithEvents objEventoCclDe As AdmEvento
Attribute objEventoCclDe.VB_VarHelpID = -1
Private WithEvents objEventoCclAte As AdmEvento
Attribute objEventoCclAte.VB_VarHelpID = -1

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 47170

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 47171

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47172
        
        ComboOpcoes.Text = ""
        DescCclInic.Caption = ""
        DescCclFim.Caption = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 47170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 47171, 47172

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170403)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47172

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 47172

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170404)

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
    If ComboOpcoes.Text = "" Then Error 47173

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47174

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 47175

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47176
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 47174, 47175, 47176

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170405)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47178
    
    ComboOpcoes.Text = ""
    DescCclInic.Caption = ""
    DescCclFim.Caption = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47178
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170406)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click

    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    If giCarregando = CANCELA Then
    
        lErro = PreencheComboPeriodo(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1)
        If lErro <> SUCESSO Then Error 47179
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 47179

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170407)

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

Dim lErro As Long
Dim objExercicio As ClassExercicio
Dim colExercicios As New Collection

On Error GoTo Erro_Form_Load

    giCarregando = CANCELA
        
    'inicializa a mascara de centro de custo/lucro inicial
    lErro = Inicializa_Mascara_CclInicial()
    If lErro <> SUCESSO Then Error 54889
    
    'inicializa a mascara de centro de custo/lucro final
    lErro = Inicializa_Mascara_CclFinal()
    If lErro <> SUCESSO Then Error 54890

'    'Inicializa a Lista de Centros de Custo
'    lErro = CF("Carga_Arvore_Ccl", TvwCcls.Nodes)
'    If lErro <> SUCESSO Then Error 47182

    Set objEventoCclDe = New AdmEvento
    Set objEventoCclAte = New AdmEvento

    '??????? Ler todos os exercicios. Trocar em todas as telas de parametros da contabilidade(confirmar com Mário antes de alterar).
    'ler os exercicios
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then Error 47184

    'preenche a Combo de exercicio
    For Each objExercicio In colExercicios
        
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
        
        If ComboExercicio.ItemData(ComboExercicio.NewIndex) = giExercicioAtual Then
            ComboExercicio.ListIndex = ComboExercicio.NewIndex
        End If
    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 47182, 47184, 54889, 54890
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170408)

    End Select

    Unload Me

    Exit Sub
    
End Sub

Private Function Inicializa_Mascara_CclInicial() As Long
'inicializa a mascara de centro de custo/lucro /m

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_CclInicial

    'Inicializa a máscara de Centro de custo/lucro
    sMascaraCcl = String(STRING_CCL, 0)
    
    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 54891
    
    'coloca a mascara na tela.
    CclInicial.Mask = sMascaraCcl
    
    Inicializa_Mascara_CclInicial = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_CclInicial:

    Inicializa_Mascara_CclInicial = Err
    
    Select Case Err
    
        Case 54891
            lErro = Rotina_Erro(vbOKOnly, "Erro_MascaraCcl", Err)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170409)
        
    End Select

    Exit Function
    
End Function

Private Function Inicializa_Mascara_CclFinal() As Long
'inicializa a mascara de centro de custo/lucro /m

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_CclFinal

    'Inicializa a máscara de Centro de custo/lucro
    sMascaraCcl = String(STRING_CCL, 0)
    
    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 54892
    
    'coloca a mascara na tela.
    CclFinal.Mask = sMascaraCcl
    
    Inicializa_Mascara_CclFinal = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_CclFinal:

    Inicializa_Mascara_CclFinal = Err
    
    Select Case Err
    
        Case 54892
            lErro = Rotina_Erro(vbOKOnly, "Erro_MascaraCcl", Err)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170410)
        
    End Select

    Exit Function
    
End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCcl_I As String
Dim sCcl_F As String
Dim iCclPreenchida_I As Integer
Dim iCclPreenchida_F As Integer

On Error GoTo Erro_PreencherRelOp

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 47185

    'período inicial não pode ser vazio
    If ComboPeriodo.Text = "" Then Error 47186
    
    'verifica se o Ccl Inicial é maior que o Ccl Final
    lErro = CF("Ccl_Formata", CclInicial.Text, sCcl_I, iCclPreenchida_I)
    If lErro <> SUCESSO Then Error 47187

    lErro = CF("Ccl_Formata", CclFinal.Text, sCcl_F, iCclPreenchida_F)
    If lErro <> SUCESSO Then Error 47188

    If (iCclPreenchida_I = CCL_PREENCHIDA) And (iCclPreenchida_F = CCL_PREENCHIDA) Then
    
        If sCcl_I > sCcl_F Then Error 47189
    
    End If
 
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 47190

    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then Error 47191
    
    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then Error 47192
    
    lErro = objRelOpcoes.IncluirParametro("NPERIODO", CStr(ComboPeriodo.ItemData(ComboPeriodo.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 47193
    
    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 47194
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47195

    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", ComboPeriodo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47196

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCcl_I, iCclPreenchida_I, sCcl_F, iCclPreenchida_F)
    If lErro <> SUCESSO Then Error 47197

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 47185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 47186
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)
            ComboPeriodo.SetFocus
                
        Case 47187, 47188
        
        Case 47189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", Err)
            
        Case 47190, 47191, 47192, 47193, 47194, 47195, 47196, 47197

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170411)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCcl_I As String, iCclPreenchida_I As Integer, sCcl_F As String, iCclPreenchida_F As Integer) As Long
'monta a expressão de seleção
'recebe as contas inicial e final no formato do BD

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If iCclPreenchida_I = CCL_PREENCHIDA Then sExpressao = "Ccl >= " & Forprint_ConvTexto(sCcl_I)

    If iCclPreenchida_F = CCL_PREENCHIDA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl <= " & Forprint_ConvTexto(sCcl_F)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170412)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim sParam As String
Dim sDescCcl As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 47198

    'pega Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro <> SUCESSO Then Error 47199
    
    If sParam <> "" Then
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 47200
    End If
    
    CclInicial.PromptInclude = False
    CclInicial.Text = sParam
    CclInicial.PromptInclude = True
        
    DescCclInic.Caption = sDescCcl
    
    'pega Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro <> SUCESSO Then Error 47201

    If sParam <> "" Then
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 47202
    End If
    
    CclFinal.PromptInclude = False
    CclFinal.Text = sParam
    CclFinal.PromptInclude = True
        
    DescCclFim.Caption = sDescCcl

    'período inicial
    lErro = objRelOpcoes.ObterParametro("NPERIODO", sParam)
    If lErro <> SUCESSO Then Error 47203

    iPeriodo = CInt(sParam)
        
    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 47204

    iExercicio = CInt(sParam)
    
    'Traz a tela o periodo e o Exercicio da opcao desejada
    lErro = MostraExercicioPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 47205
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 47198, 47199, 47200, 47201, 47202, 47203, 47204, 47205
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170413)

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
    If lErro <> SUCESSO Then Error 47206

    MostraExercicioPeriodo = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodo:

    MostraExercicioPeriodo = Err

    Select Case Err

        Case 47206

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170414)

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
    If lErro <> SUCESSO Then Error 47207
    
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

        Case 47207

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170415)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 24978
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 47183

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 47183
        
        Case 24978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170416)

    End Select

    Exit Function

End Function

Private Sub CclFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclAuxFim As String
Dim sCclFormatado As String
Dim iCclPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclFinal_Validate
    
    TvwCcls.Tag = "Ccl_Final"
    
    sCclAuxFim = CclFinal.Text

    giFocoInicial = 0
    
    If Len(CclFinal.ClipText) > 0 Then

        sCclFormatado = String(STRING_CCL, 0)

        'critica o formato do ccl e sua presença no BD
        lErro = Ccl_Critica1(CclFinal.Text, sCclFormatado, objCcl)
        If lErro <> SUCESSO And lErro <> 87164 Then gError 87185
    
        'se o centro de custo/lucro não estiver cadastrado
        If lErro = 87164 Then gError 87186

        lErro = Ccl_Perde_Foco(CclFinal, DescCclFim, objCcl)
        If lErro <> SUCESSO Then gError 81187

    End If
    
    Exit Sub
    
Erro_CclFinal_Validate:

    Cancel = True


    Select Case gErr
        
        Case 87186
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CclFinal.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatado
                
                Call Chama_Tela("CclTela", objCcl)
                
            
            Else
                
            End If

        Case 87185, 87187
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170417)
        
    End Select

    Exit Sub
    
End Sub

Private Sub CclInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sCclAuxInic As String
Dim sCclFormatado As String
Dim iCclPreenchido As Integer
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CclInicial_Validate
    
    TvwCcls.Tag = "Ccl_Inicial"
    
    sCclAuxInic = CclInicial.Text

    giFocoInicial = 1
    
    If Len(CclInicial.ClipText) > 0 Then

        sCclFormatado = String(STRING_CCL, 0)
    
        'critica o formato do ccl e sua presença no BD
        lErro = Ccl_Critica1(CclInicial.Text, sCclFormatado, objCcl) 'Analitico
        If lErro <> SUCESSO And lErro <> 87164 Then gError 87182
    
        'se o centro de custo/lucro não estiver cadastrado
        If lErro = 87164 Then gError 87183

        lErro = Ccl_Perde_Foco(CclInicial, DescCclInic, objCcl)
        If lErro <> SUCESSO Then gError 87184

    End If
        
    Exit Sub
    
Erro_CclInicial_Validate:

    Cancel = True


    Select Case gErr
            
        Case 87183
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CclInicial.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatado
                
                Call Chama_Tela("CclTela", objCcl)
                        
            End If

        Case 87182, 87184
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170418)
        
    End Select

    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCclDe = Nothing
    Set objEventoCclAte = Nothing
    
End Sub

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim lErro As Long
Dim sCcl As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_TvwCcls_NodeClick
    
    objCcl.sCcl = Right(Node.Key, Len(Node.Key) - 1)
    
    If giFocoInicial = 1 Then
        lErro = Ccl_Perde_Foco(CclInicial, DescCclInic, objCcl)
        If lErro <> SUCESSO Then gError 87180
    
    Else
        lErro = Ccl_Perde_Foco(CclFinal, DescCclFim, objCcl)
        If lErro <> SUCESSO Then gError 87181
    
    End If
    
    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case gErr

        Case 87180, 87181

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170419)

    End Select

    Exit Sub

End Sub


'Private Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
''move os dados de centro de custo/lucro do banco de dados para a arvore colNodes. /m
'
'Dim objNode As Node
'Dim colCcl As New Collection
'Dim objCcl As ClassCcl
'Dim lErro As Long
'Dim sCclMascarado As String
'Dim sCcl As String
'Dim sCclPai As String
'
'On Error GoTo Erro_Carga_Arvore_Ccl
'
'    'leitura dos centro de custo/lucro no BD
'    lErro = CF("Ccl_Le_Todos", colCcl)
'    If lErro <> SUCESSO Then Error 47215
'
'    'para cada centro de custo encontrado no bd
'    For Each objCcl In colCcl
'
'        sCclMascarado = String(STRING_CCL, 0)
'
'        'coloca a mascara no centro de custo
'        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
'        If lErro <> SUCESSO Then Error 47216
'
'        sCcl = "C" & objCcl.sCcl
'
'        sCclPai = String(STRING_CCL, 0)
'
'        'retorna o centro de custo/lucro "pai" da centro de custo/lucro em questão, se houver
'        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
'        If lErro <> SUCESSO Then Error 54704
'
'        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
'        If Len(Trim(sCclPai)) > 0 Then
'
'            sCclPai = "C" & sCclPai
'
'            'adiciona o centro de custo como filho do centro de custo pai
'            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)
'
'        Else
'
'            'se o centro de custo/lucro não possui centro de custo/lucro "pai", adiciona na árvore sem pai
'            Set objNode = colNodes.Add(, tvwLast, sCcl)
'
'        End If
'
'        'coloca o texto do nó que acabou de ser inserido
'        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
'
'    Next
'
'    Carga_Arvore_Ccl = SUCESSO
'
'    Exit Function
'
'Erro_Carga_Arvore_Ccl:
'
'    Carga_Arvore_Ccl = Err
'
'    Select Case Err
'
'        Case 54704
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)
'
'        Case 47215
'
'        Case 47216
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170420)
'
'    End Select
'
'    Exit Function
'
'End Function

''Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'''move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.
''
''Dim objNode As Node
''Dim colCcl As New Collection
''Dim objCcl As ClassCcl
''Dim lErro As Long
''Dim sCclMascarado As String
''
''On Error GoTo Erro_Carga_Arvore_Ccl
''
''    lErro = CF("Ccl_Le_Todos",colCcl)
''    If lErro <> SUCESSO Then Error 47215
''
''    For Each objCcl In colCcl
''
''        sCclMascarado = String(STRING_CCL, 0)
''
''        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
''        If lErro <> SUCESSO Then Error 47216
''
''        Set objNode = colNodes.Add(, , "C" & objCcl.sCcl, sCclMascarado & SEPARADOR & objCcl.sDescCcl)
''
''    Next
''
''    Carga_Arvore_Ccl = SUCESSO
''
''    Exit Function
''
''Erro_Carga_Arvore_Ccl:
''
''    Carga_Arvore_Ccl = Err
''
''    Select Case Err
''
''        Case 47215
''
''        Case 47216
''            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170421)
''
''    End Select
''
''    Exit Function
''
''End Function

Function Obtem_Descricao_Ccl(sCcl As String, sDescCcl As String) As Long
'recebe em sCcl o Ccl no formato do Bd
'retorna em sDescCcl a descrição do Ccl ( que será formatado para tela )

Dim lErro As Long, iCclPreenchida As Integer
Dim objCcl As New ClassCcl
Dim sCopia As String

On Error GoTo Erro_Obtem_Descricao_Ccl

    sCopia = sCcl
    sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    sCcl = String(STRING_CCL_MASK, 0)

    'determina qual Ccl deve ser lido
    objCcl.sCcl = sCopia

    lErro = Mascara_MascararCcl(sCopia, sCcl)
    If lErro <> SUCESSO Then Error 47217

    'verifica se a conta está preenchida
    lErro = CF("Ccl_Formata", sCcl, sCopia, iCclPreenchida)
    If lErro <> SUCESSO Then Error 47218

    If iCclPreenchida = CCL_PREENCHIDA Then

        'verifica se a Ccl existe
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO Then Error 47219

        sDescCcl = objCcl.sDescCcl

    Else

        sCcl = ""
        sDescCcl = ""

    End If

    Obtem_Descricao_Ccl = SUCESSO

    Exit Function

Erro_Obtem_Descricao_Ccl:

    Obtem_Descricao_Ccl = Err

    Select Case Err

        Case 47217
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCopia)

        Case 47218, 47219

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170422)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ORC_REAL_CCL
    Set Form_Load_Ocx = Me
    Caption = "Realizado x Orçado por Centro de Custo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpOrcRealCcl"
    
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



'Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label5, Source, X, Y)
'End Sub
'
'Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label4, Source, X, Y)
'End Sub
'
'Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
'End Sub

Private Sub DescCclInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCclInic, Source, X, Y)
End Sub

Private Sub DescCclInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCclInic, Button, Shift, X, Y)
End Sub

Private Sub DescCclFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCclFim, Source, X, Y)
End Sub

Private Sub DescCclFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCclFim, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
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

'??? Fernando subrir Função
Function Ccl_Perde_Foco(objCclCod As Object, objDescCcl As Object, objCcl As ClassCcl) As Long

Dim sCclEnxuta As String
Dim lErro As Long
Dim lPosicaoSeparador As Long
Dim sCcl As String
    
On Error GoTo Erro_Ccl_Perde_Foco
    
    sCcl = objCcl.sCcl
        
    sCclEnxuta = String(STRING_CCL, 0)
    
    'volta mascarado apenas os caracteres preenchidos
    lErro = Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
    If lErro <> SUCESSO Then gError 87158

    'Preenche a Ccl com o código mascarado
    objCclCod.PromptInclude = False
    objCclCod.Text = sCclEnxuta
    objCclCod.PromptInclude = True

    
    'Faz leitura na tabela afim de saber a descrição
    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO Then gError 87169
    
    'Preenche a descrição da Ccl
    objDescCcl.Caption = objCcl.sDescCcl

    Exit Function

Erro_Ccl_Perde_Foco:

    Select Case gErr

        Case 87158
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, sCcl)

        Case 87169

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170423)

    End Select

    Exit Function

End Function

'??? Fernando subir função
Function Ccl_Critica1(ByVal sCcl As String, sCclFormatada As String, objCcl As ClassCcl) As Long
'critica o formato do ccl e sua presença no BD


Dim lErro As Long
Dim iCclPreenchida As Integer

On Error GoTo Erro_Ccl_Critica1

    If Len(sCcl) > 0 Then
    
        lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 87162
    
        If iCclPreenchida = CCL_PREENCHIDA Then
        
            objCcl.sCcl = sCclFormatada
    
            lErro = CF("Ccl_Le", objCcl)
            If lErro <> SUCESSO And lErro <> 5599 Then gError 87163
    
            'Ausencia de Ccl no BD
            If lErro = 5599 Then gError 87164
                        
        End If
        
    End If
    
    Ccl_Critica1 = SUCESSO
    
    Exit Function

Erro_Ccl_Critica1:

    Ccl_Critica1 = gErr
    
    Select Case gErr
    
        Case 87162, 87163, 87164
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170424)
        
    End Select
    
    Exit Function

End Function

Public Sub LabelCclDe_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
Dim sCclOrigem As String
Dim iCclPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_LabelCclDe_Click

    If Len(Trim(CclInicial.ClipText)) > 0 Then
    
        lErro = CF("Ccl_Formata", CclInicial.Text, sCclOrigem, iCclPreenchida)
        If lErro <> SUCESSO Then gError 197943

        If iCclPreenchida = CCL_PREENCHIDA Then objCcl.sCcl = sCclOrigem
    Else
        objCcl.sCcl = ""
    End If

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclDe)
    
    Exit Sub
    
Erro_LabelCclDe_Click:

    Select Case gErr
        
        Case 197943
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197945)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoCclDe_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclEnxuta As String

On Error GoTo Erro_objEventoCclDe_evSelecao
    
    Set objCcl = obj1

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
    If lErro <> SUCESSO Then gError 197947

    CclInicial.PromptInclude = False
    CclInicial.Text = sCclEnxuta
    CclInicial.PromptInclude = True
    Call CclInicial_Validate(bSGECancelDummy)

    Me.Show
    
    Exit Sub
    
Erro_objEventoCclDe_evSelecao:

    Select Case gErr

        Case 197947
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197948)
        
    End Select

    Exit Sub

End Sub

Public Sub LabelCclAte_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
Dim sCclOrigem As String
Dim iCclPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_LabelCclAte_Click

    If Len(Trim(CclFinal.ClipText)) > 0 Then
    
        lErro = CF("Ccl_Formata", CclFinal.Text, sCclOrigem, iCclPreenchida)
        If lErro <> SUCESSO Then gError 197943

        If iCclPreenchida = CCL_PREENCHIDA Then objCcl.sCcl = sCclOrigem
    Else
        objCcl.sCcl = ""
    End If

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclAte)
    
    Exit Sub
    
Erro_LabelCclAte_Click:

    Select Case gErr
        
        Case 197943
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197945)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoCclAte_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclEnxuta As String

On Error GoTo Erro_objEventoCclAte_evSelecao
    
    Set objCcl = obj1

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
    If lErro <> SUCESSO Then gError 197947

    CclFinal.PromptInclude = False
    CclFinal.Text = sCclEnxuta
    CclFinal.PromptInclude = True
    Call CclFinal_Validate(bSGECancelDummy)

    Me.Show
    
    Exit Sub
    
Erro_objEventoCclAte_evSelecao:

    Select Case gErr

        Case 197947
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197948)
        
    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CclFinal Then Call LabelCclAte_Click
        If Me.ActiveControl Is CclInicial Then Call LabelCclDe_Click
    
    End If
    
End Sub

