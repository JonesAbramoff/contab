VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpDemResPerCclOcx 
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   8805
   Begin VB.Frame Frame2 
      Caption         =   "Centro de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   225
      TabIndex        =   20
      Top             =   2520
      Width           =   8400
      Begin MSMask.MaskEdBox CclInicial 
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Top             =   420
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CclFinal 
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Top             =   1020
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label LabelcclDe 
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
         Height          =   255
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   435
         Width           =   735
      End
      Begin VB.Label LabelCclAte 
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
         Height          =   255
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label DescCclInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2835
         TabIndex        =   24
         Top             =   405
         Width           =   5415
      End
      Begin VB.Label DescCclFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2835
         TabIndex        =   23
         Top             =   1020
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Diário Geral"
      Height          =   795
      Left            =   8295
      TabIndex        =   15
      Top             =   4065
      Visible         =   0   'False
      Width           =   5715
      Begin VB.TextBox PrimeiraFolha 
         Height          =   285
         Left            =   4860
         TabIndex        =   17
         Top             =   345
         Width           =   510
      End
      Begin VB.TextBox Diario 
         Height          =   285
         Left            =   1740
         TabIndex        =   16
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número da Primeira Folha:"
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
         Left            =   2580
         TabIndex        =   19
         Top             =   375
         Width           =   2250
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número do Diário:"
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
         Left            =   180
         TabIndex        =   18
         Top             =   345
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6480
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDemResPerCclOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDemResPerCclOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDemResPerCclOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDemResPerCclOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDemResPerCclOcx.ctx":0994
      Left            =   1050
      List            =   "RelOpDemResPerCclOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpDemResPerCclOcx.ctx":0998
      Left            =   1050
      List            =   "RelOpDemResPerCclOcx.ctx":099A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   750
      Width           =   1860
   End
   Begin VB.ComboBox ComboPeriodo 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
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
      Height          =   720
      Left            =   4455
      Picture         =   "RelOpDemResPerCclOcx.ctx":099C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   135
      Width           =   1455
   End
   Begin VB.ComboBox ComboModelos 
      Height          =   315
      ItemData        =   "RelOpDemResPerCclOcx.ctx":0A9E
      Left            =   1020
      List            =   "RelOpDemResPerCclOcx.ctx":0AA0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1875
      Width           =   2775
   End
   Begin VB.CommandButton BotaoConfigura 
      Caption         =   "Configurar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4455
      Picture         =   "RelOpDemResPerCclOcx.ctx":0AA2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1620
      Width           =   1455
   End
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   2970
      Left            =   6030
      TabIndex        =   27
      Top             =   1050
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   5239
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Centros de Custo / Lucro"
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
      Left            =   6060
      TabIndex        =   28
      Top             =   825
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   285
      Width           =   615
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1350
      Width           =   735
   End
   Begin VB.Label Label6 
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
      TabIndex        =   7
      Top             =   810
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
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
      Left            =   285
      TabIndex        =   6
      Top             =   1920
      Width           =   690
   End
End
Attribute VB_Name = "RelOpDemResPerCclOcx"
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

Private WithEvents objEventoCclDe As AdmEvento
Attribute objEventoCclDe.VB_VarHelpID = -1
Private WithEvents objEventoCclAte As AdmEvento
Attribute objEventoCclAte.VB_VarHelpID = -1

Function MostraExercicioPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama PreencheComboPeriodo

Dim iConta As Integer, lErro As Long

On Error GoTo Erro_MostraExercicioPeriodo
    
    'seta cargando para que não execute o click
    giCarregando = OK

    For iConta = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iConta) = iExercicio Then
            ComboExercicio.ListIndex = iConta
            Exit For
        End If
    Next

    lErro = PreencheComboPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 13390

    MostraExercicioPeriodo = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodo:

    MostraExercicioPeriodo = Err

    Select Case Err

        Case 13390

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168174)

    End Select

    Exit Function

End Function

Function PreencheComboPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'lê os períodos do exercício 'iExercicio' preenchendo o combo de período
'seleciona o período 'iPeriodo'

Dim lErro As Long, iConta As Integer
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheComboPeriodo

    ComboPeriodo.Clear

    'inicializar os periodos do exercicio selecionado no combo de exercícios
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 13391

    For iConta = 1 To colPeriodos.Count
        Set objPeriodo = colPeriodos.Item(iConta)
        ComboPeriodo.AddItem objPeriodo.sNomeExterno
        ComboPeriodo.ItemData(ComboPeriodo.NewIndex) = objPeriodo.iPeriodo
    Next

    'mostra o período
    For iConta = 0 To ComboPeriodo.ListCount - 1
        If ComboPeriodo.ItemData(iConta) = iPeriodo Then
            ComboPeriodo.ListIndex = iConta
            Exit For
        End If
    Next

    PreencheComboPeriodo = SUCESSO

    Exit Function

Erro_PreencheComboPeriodo:

    PreencheComboPeriodo = Err

    Select Case Err

        Case 13391

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168175)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29592
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 13411
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 13411
        
        Case 29592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168176)

    End Select

    Exit Function

End Function

Private Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional sCcl_I As String, Optional sCcl_F As String) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lCodIdentExec As Long
Dim iCclPreenchida_I As Integer, iCclPreenchida_F As Integer

On Error GoTo Erro_PreencherRelOp

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 13404

    'período não pode ser vazio
    If ComboPeriodo.Text = "" Then Error 13405

    'verifica se o Ccl Inicial é maior que o Ccl Final
    lErro = CF("Ccl_Formata", CclInicial.Text, sCcl_I, iCclPreenchida_I)
    If lErro <> SUCESSO Then Error 13477

    lErro = CF("Ccl_Formata", CclFinal.Text, sCcl_F, iCclPreenchida_F)
    If lErro <> SUCESSO Then Error 13478

    If (iCclPreenchida_I = CCL_PREENCHIDA) And (iCclPreenchida_F = CCL_PREENCHIDA) Then
    
        If sCcl_I > sCcl_F Then Error 13479
    
    End If
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 13393

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13394
    
    lErro = objRelOpcoes.IncluirParametro("TEXERCICIO", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 59516

    lErro = objRelOpcoes.IncluirParametro("NPERIODO", CStr(ComboPeriodo.ItemData(ComboPeriodo.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13395
    
    lErro = objRelOpcoes.IncluirParametro("TPERIODO", ComboPeriodo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 59517
    
    lErro = objRelOpcoes.IncluirParametro("TMODELO", ComboModelos.Text)
    If lErro <> AD_BOOL_TRUE Then Error 59518

    lErro = objRelOpcoes.IncluirParametro("NPAGRELINI", PrimeiraFolha.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13182

    lErro = objRelOpcoes.IncluirParametro("NNUMDIARIO", Diario.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13184

    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then Error 13481

    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then Error 13482

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err
    
    Select Case Err

        Case 7115, 59516, 59517, 59518, 13182, 13184
        
        Case 13404
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 13405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)
            ComboPeriodo.SetFocus

        Case 13393

        Case 13394, 13395

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168177)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long, iExercicio As Integer, iPeriodo As Integer
Dim sParam As String
Dim sDescCcl As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 13396

    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 13397

    iExercicio = CInt(sParam)

    'período
    lErro = objRelOpcoes.ObterParametro("NPERIODO", sParam)
    If lErro <> SUCESSO Then Error 13398

    iPeriodo = CInt(sParam)

    lErro = MostraExercicioPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 13399
    
    'modelo
    lErro = objRelOpcoes.ObterParametro("TMODELO", sParam)
    If lErro <> SUCESSO Then Error 59519
    
    ComboModelos.Text = sParam

    'pega primeira folha e exibe
    lErro = objRelOpcoes.ObterParametro("NPAGRELINI", sParam)
    If lErro <> SUCESSO Then Error 13188

    PrimeiraFolha.Text = sParam

    'pega número diário e exibe
    lErro = objRelOpcoes.ObterParametro("NNUMDIARIO", sParam)
    If lErro <> SUCESSO Then Error 13188

    Diario.Text = sParam

    'pega Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro <> SUCESSO Then Error 13467

    If sParam <> "" Then
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 13468
    End If
    
    CclInicial.PromptInclude = False
    CclInicial.Text = sParam
    CclInicial.PromptInclude = True
    
    DescCclInic.Caption = sDescCcl
    
    'pega Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro <> SUCESSO Then Error 13469

    If sParam <> "" Then
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 13470
    End If
    
    CclFinal.PromptInclude = False
    CclFinal.Text = sParam
    CclFinal.PromptInclude = True
    
    DescCclFim.Caption = sDescCcl

    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err
    
    Select Case Err

        Case 13396, 13397, 13398, 13399, 59519, 13188

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168178)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoConfigura_Click()

Dim iIndice As Integer
Dim lErro As Long
Dim colModelos As New Collection

On Error GoTo Erro_BotaoConfigura_Click

    Call Chama_Tela("RelDREConfig", RELDRE, "Configuração do Demonstrativo de Resultados")
    
    Exit Sub

Erro_BotaoConfigura_Click:
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168179)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 13400

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 13401

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47044
    
        DescCclInic.Caption = ""
        DescCclFim.Caption = ""
    
    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 13400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 13401, 47044

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168180)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim lNumIntRel As Long
Dim sCcl_I As String, sCcl_F As String

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, sCcl_I, sCcl_F)
    If lErro <> SUCESSO Then Error 13402

    lErro = RelDRPCcl_Calcula(ComboModelos.List(ComboModelos.ListIndex), ComboExercicio.ItemData(ComboExercicio.ListIndex), ComboPeriodo.ItemData(ComboPeriodo.ListIndex), giFilialEmpresa, sCcl_I, sCcl_F, lNumIntRel)
    If lErro <> SUCESSO Then Error 43586
        
    lErro = gobjRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
    If lErro <> AD_BOOL_TRUE Then Error 7121
        
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err
        Case 7116
        Case 13402

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168181)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function RelDRPCcl_Calcula(ByVal sModelo As String, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal iFilialEmpresa As Integer, ByVal sCclInicial As String, ByVal sCclFinal As String, lNumIntRel As Long) As Long
'Calcula Valor correspondente ao Modelo

Dim lErro As Long
Dim alComando(0 To 3) As Long, iIndice As Integer
Dim lTransacao As Long, sAux As String, sCcl As String

On Error GoTo Erro_RelDRPCcl_Calcula
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184107
    Next
    
    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184108
    
    'obtem numintrel
    lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_REL_DRP", lNumIntRel)
    If lErro <> SUCESSO Then gError 184109
    
    sCcl = String(STRING_CCL, 0)
    sAux = Format(iPeriodo, "00")
    
    If Len(Trim(sCclInicial)) <> 0 And Len(Trim(sCclFinal)) <> 0 Then
        lErro = Comando_Executar(alComando(0), "SELECT DISTINCT MvPerCcl.Ccl FROM MvPerCcl, Ccl WHERE MvPerCcl.Ccl = Ccl.Ccl AND AtivoCcl = 1 AND TipoCcl = 3 AND FilialEmpresa = ? AND Exercicio = ? AND (Cre" & sAux & "<>0 OR Deb" & sAux & "<>0) AND MvPerCcl.Ccl BETWEEN ? AND ? ORDER BY MvPerCcl.Ccl", sCcl, iFilialEmpresa, iExercicio, sCclInicial, sCclFinal)
    Else
        If Len(Trim(sCclInicial)) = 0 And Len(Trim(sCclFinal)) = 0 Then
            lErro = Comando_Executar(alComando(0), "SELECT DISTINCT MvPerCcl.Ccl FROM MvPerCcl, Ccl WHERE MvPerCcl.Ccl = Ccl.Ccl AND AtivoCcl = 1 AND TipoCcl = 3 AND FilialEmpresa = ? AND Exercicio = ? AND (Cre" & sAux & "<>0 OR Deb" & sAux & "<>0) ORDER BY MvPerCcl.Ccl", sCcl, iFilialEmpresa, iExercicio)
        Else
            If Len(Trim(sCclInicial)) <> 0 Then
                lErro = Comando_Executar(alComando(0), "SELECT DISTINCT MvPerCcl.Ccl FROM MvPerCcl, Ccl WHERE MvPerCcl.Ccl = Ccl.Ccl AND AtivoCcl = 1 AND TipoCcl = 3 AND FilialEmpresa = ? AND Exercicio = ? AND (Cre" & sAux & "<>0 OR Deb" & sAux & "<>0) AND MvPerCcl.Ccl >= ? ORDER BY MvPerCcl.Ccl", sCcl, iFilialEmpresa, iExercicio, sCclInicial)
            Else
                lErro = Comando_Executar(alComando(0), "SELECT DISTINCT MvPerCcl.Ccl FROM MvPerCcl, Ccl WHERE MvPerCcl.Ccl = Ccl.Ccl AND AtivoCcl = 1 AND TipoCcl = 3 AND FilialEmpresa = ? AND Exercicio = ? AND (Cre" & sAux & "<>0 OR Deb" & sAux & "<>0) AND MvPerCcl.Ccl <= ? ORDER BY MvPerCcl.Ccl", sCcl, iFilialEmpresa, iExercicio, sCclFinal)
            End If
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 184110
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184111
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = RelDRPCcl_Calcula1(sModelo, iExercicio, iPeriodo, iFilialEmpresa, sCcl, lNumIntRel, alComando)
        If lErro <> SUCESSO Then gError 184112
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184113
    
    Loop
    
    'Confirma a transção
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184114
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelDRPCcl_Calcula = SUCESSO
    
    Exit Function
      
Erro_RelDRPCcl_Calcula:

    RelDRPCcl_Calcula = gErr
    
    Select Case gErr
    
        Case 184109, 184112
        
        Case 184110, 184111, 184113
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRPCCL", gErr)
            
        Case 184107
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 184108
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 184114
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184115)
        
    End Select
    
    'Rollback
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 13403

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13406

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 13407

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47045
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 13403
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 13406

        Case 13407, 47045
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168183)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa os campos da Tela
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47047
    
    DescCclInic.Caption = ""
    DescCclFim.Caption = ""
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47047
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168184)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click
    
    'se esta vazia
    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    'se não estiver carregando do BD
    If giCarregando = CANCELA Then
        
        'preenche a combo com periodo 1
        lErro = PreencheComboPeriodo(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1)
        If lErro <> SUCESSO Then Error 13408
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 13408

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168185)

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
Dim colExerciciosAbertos As New Collection
Dim colModelos As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    giCarregando = CANCELA

    'inicializa a mascara de centro de custo/lucro inicial
    lErro = Inicializa_Mascara_CclInicial()
    If lErro <> SUCESSO Then Error 54877
    
    'inicializa a mascara de centro de custo/lucro final
    lErro = Inicializa_Mascara_CclFinal()
    If lErro <> SUCESSO Then Error 54878

'    'Inicializa a Lista de Centros de Custo
'    lErro = CF("Carga_Arvore_Ccl", TvwCcls.Nodes)
'    If lErro <> SUCESSO Then Error 13502

    Set objEventoCclDe = New AdmEvento
    Set objEventoCclAte = New AdmEvento

    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 13412
    
    For iIndice = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iIndice)
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next

    'Le os Modelos na tabela RelDRE
    lErro = CF("RelDRE_Le_Modelos_Distintos", RELDRE, colModelos)
    If lErro <> SUCESSO And lErro <> 47101 Then Error 47102
    
    If lErro = SUCESSO Then
    
        'preenche a combo Modelos
        For iIndice = 1 To colModelos.Count
            ComboModelos.AddItem colModelos.Item(iIndice)
        Next
        
        ComboModelos.ListIndex = -1
        
        
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 13412, 47102
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168186)

    End Select

    Unload Me

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCclDe = Nothing
    Set objEventoCclAte = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DEMONS_RESULT_PERIODO
    Set Form_Load_Ocx = Me
    Caption = "Demonstrativo de Resultado Período/Centro de Custo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDemResPerCcl"
    
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub


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
    If lErro <> SUCESSO Then Error 13460

    'verifica se a conta está preenchida
    lErro = CF("Ccl_Formata", sCcl, sCopia, iCclPreenchida)
    If lErro <> SUCESSO Then Error 13461

    If iCclPreenchida = CCL_PREENCHIDA Then

        'verifica se a Ccl existe
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO Then Error 13462

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

        Case 13460
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCopia)

        Case 13461

        Case 13462

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169612)

    End Select

    Exit Function

End Function

Private Sub CclFinal_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sCclFormatado As String
Dim iCclPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclFinal_Validate
    
    giFocoInicial = 0
    
    If Len(CclFinal.ClipText) > 0 Then

        sCclFormatado = String(STRING_CCL, 0)

        'critica o formato do ccl e sua presença no BD
        lErro = Ccl_Critica1(CclFinal.Text, sCclFormatado, objCcl)
        If lErro <> SUCESSO And lErro <> 87164 Then gError 87177
    
        'se o centro de custo/lucro não estiver cadastrado
        If lErro = 87164 Then gError 87178

        lErro = Ccl_Perde_Foco(CclFinal, DescCclFim, objCcl)
        If lErro <> SUCESSO Then gError 81179

    End If
    
    Exit Sub
    
Erro_CclFinal_Validate:

    Cancel = True


    Select Case gErr
        
        Case 87178
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CclFinal.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatado
                
                Call Chama_Tela("CclTela", objCcl)
                                
            End If

        Case 87177, 87179
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169625)
        
    End Select

    Exit Sub
    
End Sub

Private Sub CclInicial_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sCclFormatado As String
Dim iCclPreenchido As Integer
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CclInicial_Validate
    
    giFocoInicial = 1
    
    If Len(CclInicial.ClipText) > 0 Then

        sCclFormatado = String(STRING_CCL, 0)
    
        'critica o formato do ccl e sua presença no BD
        lErro = Ccl_Critica1(CclInicial.Text, sCclFormatado, objCcl) 'Analitico
        If lErro <> SUCESSO And lErro <> 87164 Then gError 87174
    
        'se o centro de custo/lucro não estiver cadastrado
        If lErro = 87164 Then gError 87175

        lErro = Ccl_Perde_Foco(CclInicial, DescCclInic, objCcl)
        If lErro <> SUCESSO Then gError 87176

    End If
        
    Exit Sub
    
Erro_CclInicial_Validate:

    Cancel = True


    Select Case gErr
            
        Case 87175
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CclInicial.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatado
                
                Call Chama_Tela("CclTela", objCcl)
                        
            End If

        Case 87174, 87176
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169626)
        
    End Select

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
    If lErro <> SUCESSO Then Error 54879
    
    'coloca a mascara na tela.
    CclInicial.Mask = sMascaraCcl
    
    Inicializa_Mascara_CclInicial = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_CclInicial:

    Inicializa_Mascara_CclInicial = Err
    
    Select Case Err
    
        Case 54879
            lErro = Rotina_Erro(vbOKOnly, "Erro_MascaraCcl", Err)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169629)
        
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
    If lErro <> SUCESSO Then Error 54880
    
    'coloca a mascara na tela.
    CclFinal.Mask = sMascaraCcl
    
    Inicializa_Mascara_CclFinal = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_CclFinal:

    Inicializa_Mascara_CclFinal = Err
    
    Select Case Err
    
        Case 54880
            lErro = Rotina_Erro(vbOKOnly, "Erro_MascaraCcl", Err)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169630)
        
    End Select

    Exit Function
    
End Function

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim lErro As Long
Dim sCcl As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_TvwCcls_NodeClick
    
    objCcl.sCcl = Right(Node.Key, Len(Node.Key) - 1)
    
    If giFocoInicial = 1 Then
        lErro = Ccl_Perde_Foco(CclInicial, DescCclInic, objCcl)
        If lErro <> SUCESSO Then gError 87172
    
    Else
        lErro = Ccl_Perde_Foco(CclFinal, DescCclFim, objCcl)
        If lErro <> SUCESSO Then gError 87173
    
    End If
    
    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case gErr

        Case 87172, 87173

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169631)

    End Select

    Exit Sub

End Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169632)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169633)
        
    End Select
    
    Exit Function

End Function

Private Function RelDRPCcl_Calcula1(ByVal sModelo As String, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal iFilialEmpresa As Integer, ByVal sCcl As String, lNumIntRel As Long, alComando() As Long) As Long
'Calcula Valor correspondente ao Modelo

Dim lErro As Long
Dim iOperacao As Integer
Dim dValorPerAnt As Double
Dim dValorPerAtual As Double
Dim dValorPerAcumAnt As Double
Dim dValorPerAcumAtual As Double
Dim colRelDRE As New Collection
Dim colRelDREConta As New Collection
Dim colRelDREFormula As New Collection
Dim objRelDRE As New ClassRelDRE
Dim objRelDRE1 As New ClassRelDRE
Dim objRelDREConta As New ClassRelDREConta
Dim objRelDREFormula As New ClassRelDREFormula
'Dim alComando(0 To 3) As Long, iIndice As Integer

On Error GoTo Erro_RelDRPCcl_Calcula1
    
    'Lê registros da tabela RelDRE para o Modelo passado como parâmetro
    lErro = CF("RelDRE_Le_Modelo", RELDRE, sModelo, colRelDRE)
    If lErro <> SUCESSO Then gError 184116
    
    'Lê registros da tabela RelDREConta para o Modelo passado como parâmetro
    lErro = CF("RelDREConta_Le_Modelo", RELDRE, sModelo, colRelDREConta)
    If lErro <> SUCESSO Then gError 184117
    
    'Lê registros da tabela RelDREFormula para o Modelo passado como parâmetro
    lErro = CF("RelDREFormula_Le_Modelo", RELDRE, sModelo, colRelDREFormula)
    If lErro <> SUCESSO Then gError 184118
    
    For Each objRelDRE In colRelDRE
        
        If objRelDRE.iTipo = DRE_TIPO_CONTA Then
        
            For Each objRelDREConta In colRelDREConta
                
                'Se o elemento tiver o mesmo código do elemento da coleção RelDRE
                If objRelDREConta.iCodigo = objRelDRE.iCodigo Then
                    
                    If objRelDRE.iExercicio = CONTAS_EXERCICIO_ANTERIOR Then
                    
                        'Calcula o Valor
                        lErro = CF("MvPerCcl_Calcula_Valor_Periodo", iFilialEmpresa, iExercicio - 1, iPeriodo, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, dValorPerAnt, dValorPerAtual, dValorPerAcumAnt, dValorPerAcumAtual, sCcl, sCcl)
                        If lErro <> SUCESSO Then gError 184119
                        
                    Else
                    
                        'Calcula o Valor
                        lErro = CF("MvPerCcl_Calcula_Valor_Periodo", iFilialEmpresa, iExercicio, iPeriodo, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, dValorPerAnt, dValorPerAtual, dValorPerAcumAnt, dValorPerAcumAtual, sCcl, sCcl)
                        If lErro <> SUCESSO Then gError 184120
                    
                    End If
                    
                    'Aculmula o Valor
                    objRelDRE.dValor = objRelDRE.dValor + dValorPerAtual
                    objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt + dValorPerAnt
                    objRelDRE.dValorPerAcumAnt = objRelDRE.dValorPerAcumAnt + dValorPerAcumAnt
                    objRelDRE.dValorPerAcumAtual = objRelDRE.dValorPerAcumAtual + dValorPerAcumAtual
                    
                End If
                
            Next
            
        ElseIf objRelDRE.iTipo = DRE_TIPO_FORMULA Then
        
            iOperacao = DRE_OPERACAO_SOMA
            
            For Each objRelDREFormula In colRelDREFormula
                
                'Se o elemento tiver o mesmo código do elemento da coleção RelDRE
                If objRelDREFormula.iCodigo = objRelDRE.iCodigo Then
                                        
                    For Each objRelDRE1 In colRelDRE
                        
                        'Pesquisar na coleção RelDRE o elemento que tenha o Código igual ao
                        'código da fórmula do elemento da coleção RelDREFormula
                        If objRelDRE1.iCodigo = objRelDREFormula.iCodigoFormula Then
                        
                            If iOperacao = DRE_OPERACAO_SOMA Then
                                objRelDRE.dValor = objRelDRE.dValor + objRelDRE1.dValor
                                objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt + objRelDRE1.dValorExercAnt
                                objRelDRE.dValorPerAcumAnt = objRelDRE.dValorPerAcumAnt + objRelDRE1.dValorPerAcumAnt
                                objRelDRE.dValorPerAcumAtual = objRelDRE.dValorPerAcumAtual + objRelDRE1.dValorPerAcumAtual
                            Else
                                objRelDRE.dValor = objRelDRE.dValor - objRelDRE1.dValor
                                objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt - objRelDRE1.dValorExercAnt
                                objRelDRE.dValorPerAcumAnt = objRelDRE.dValorPerAcumAnt - objRelDRE1.dValorPerAcumAnt
                                objRelDRE.dValorPerAcumAtual = objRelDRE.dValorPerAcumAtual - objRelDRE1.dValorPerAcumAtual
                            End If
                            
                            iOperacao = objRelDREFormula.iOperacao
                
                            Exit For
                            
                        End If
                        
                    Next
                        
                End If
                
            Next
        
        End If
        
    Next
        
    For Each objRelDRE In colRelDRE
    
        If objRelDRE.iImprime <> 0 Then
                        
            If objRelDRE.dValor <> 0 Or objRelDRE.dValorExercAnt <> 0 Or objRelDRE.dValorPerAcumAnt <> 0 Or objRelDRE.dValorPerAcumAtual <> 0 Then
            
                lErro = Comando_Executar(alComando(1), "INSERT INTO RelDRECcl (NumIntRel, Relatorio, Modelo, Codigo, ContaInicial, ContaFinal, CclInicial, CclFinal, Ccl, Tipo, Nivel, Titulo, Posicao, Valor, ValorExercAnt, ValorPerAcumAnt, ValorPerAcumAtual, Exercicio) " & _
                    " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    lNumIntRel, RELDRE, sModelo, objRelDRE.iCodigo, "", "", "", "", sCcl, objRelDRE.iTipo, objRelDRE.iNivel, objRelDRE.sTitulo, objRelDRE.iPosicao, objRelDRE.dValor, objRelDRE.dValorExercAnt, objRelDRE.dValorPerAcumAnt, objRelDRE.dValorPerAcumAtual, iExercicio)
                If lErro <> AD_SQL_SUCESSO Then gError 184121
    
            End If
            
        End If
        
    Next
        
    RelDRPCcl_Calcula1 = SUCESSO
    
    Exit Function
      
Erro_RelDRPCcl_Calcula1:

    RelDRPCcl_Calcula1 = gErr
    
    Select Case gErr
    
        Case 184116, 184117, 184118, 184119, 184120
        
        Case 184121
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_RELDRECCL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184122)
        
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
