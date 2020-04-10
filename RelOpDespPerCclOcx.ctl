VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpDespPerCclOcx 
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   8550
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6240
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDespPerCclOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDespPerCclOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDespPerCclOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDespPerCclOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   12
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
      Left            =   4260
      Picture         =   "RelOpDespPerCclOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDespPerCclOcx.ctx":0A96
      Left            =   1110
      List            =   "RelOpDespPerCclOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2655
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpDespPerCclOcx.ctx":0A9A
      Left            =   1110
      List            =   "RelOpDespPerCclOcx.ctx":0A9C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1095
      Width           =   1620
   End
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
      Height          =   1500
      Left            =   120
      TabIndex        =   16
      Top             =   1650
      Width           =   8235
      Begin MSMask.MaskEdBox CclInicial 
         Height          =   285
         Left            =   930
         TabIndex        =   4
         Top             =   360
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CclFinal 
         Height          =   285
         Left            =   930
         TabIndex        =   5
         Top             =   975
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label LabelCclDe 
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
         Left            =   270
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   375
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   990
         Width           =   615
      End
      Begin VB.Label DescCclInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2940
         TabIndex        =   18
         Top             =   360
         Width           =   5100
      End
      Begin VB.Label DescCclFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2940
         TabIndex        =   17
         Top             =   975
         Width           =   5100
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Periodos"
      Height          =   645
      Left            =   2820
      TabIndex        =   13
      Top             =   870
      Width           =   5535
      Begin VB.ComboBox PeriodoInicial 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   1695
      End
      Begin VB.ComboBox PeriodoFinal 
         Height          =   315
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   285
         TabIndex        =   15
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3150
         TabIndex        =   14
         Top             =   270
         Width           =   480
      End
   End
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   2535
      Left            =   6015
      TabIndex        =   6
      Top             =   1005
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   4471
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
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
      Left            =   435
      TabIndex        =   23
      Top             =   330
      Width           =   615
   End
   Begin VB.Label Label6 
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
      Left            =   210
      TabIndex        =   22
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   6015
      TabIndex        =   21
      Top             =   765
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "RelOpDespPerCclOcx"
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

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCclDe = Nothing
    Set objEventoCclAte = Nothing
    
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
    If lErro <> SUCESSO Then Error 13458

    'verifica se a conta está preenchida
    lErro = CF("Ccl_Formata", sCcl, sCopia, iCclPreenchida)
    If lErro <> SUCESSO Then Error 13459

    If iCclPreenchida = CCL_PREENCHIDA Then

        'verifica se a Ccl existe
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO Then Error 13457

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

        Case 13457

        Case 13458
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCopia)

        Case 13459

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168208)

    End Select

    Exit Function

End Function


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCcl_I As String, iCclPreenchida_I As Integer, sCcl_F As String, iCclPreenchida_F As Integer) As Long
'monta a expressão de seleção
'recebe os ccl's inicial e final no formato do BD

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If iCclPreenchida_I = CCL_PREENCHIDA Then sExpressao = "Ccl >= " & Forprint_ConvTexto(sCcl_I)

    If iCclPreenchida_F = CCL_PREENCHIDA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl <= " & Forprint_ConvTexto(sCcl_F)

    End If

'    If sExpressao <> "" Then sExpressao = sExpressao & " E "
'    Select Case giFilialEmpresa
'        Case EMPRESA_TODA
'            sExpressao = sExpressao & "FilialEmpresa < " & Forprint_ConvInt(Abs(giFilialAuxiliar))
'
'        Case Abs(giFilialAuxiliar)
'            sExpressao = sExpressao & "FilialEmpresa > " & Forprint_ConvInt(Abs(giFilialAuxiliar))
'
'        Case Else
'            sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
'    End Select
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168209)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long, iExercicio As Integer, iPeriodo As Integer
Dim sParam As String
Dim sDescCcl As String
Dim iPer_I  As Integer
Dim iPer_F As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 13417

    'pega Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro <> SUCESSO Then Error 13418
    
    If sParam <> "" Then
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 13419
    End If
    
    CclInicial.PromptInclude = False
    CclInicial.Text = sParam
    CclInicial.PromptInclude = True
    
    DescCclInic.Caption = sDescCcl
    
    'pega Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro <> SUCESSO Then Error 13420

    If sParam <> "" Then
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 13421
    End If
    
    CclFinal.PromptInclude = False
    CclFinal.Text = sParam
    CclFinal.PromptInclude = True
    
    DescCclFim.Caption = sDescCcl
    
    'período inicial
    lErro = objRelOpcoes.ObterParametro("NPERIODOINICIAL", sParam)
    If lErro <> SUCESSO Then Error 13422

    iPer_I = CInt(sParam)

    'período final
    lErro = objRelOpcoes.ObterParametro("NPERIODO", sParam)
    If lErro <> SUCESSO Then Error 19399

    iPer_F = CInt(sParam)

    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 13423

    iExercicio = CInt(sParam)

    lErro = MostraExercicioPeriodos(iExercicio, iPer_I, iPer_F)
    If lErro <> SUCESSO Then Error 13424
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 13417, 13418, 13420, 13422, 13423, 13419, 13421, 13424, 19399

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168210)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCcl_I As String, sCcl_F As String
Dim iCclPreenchida_I As Integer, iCclPreenchida_F As Integer

On Error GoTo Erro_PreencherRelOp

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 13426

    'período não pode ser vazio
    If PeriodoInicial.Text = "" Then Error 13427
    If PeriodoFinal.Text = "" Then Error 47169

    'verifica se o Ccl Inicial é maior que o Ccl Final
    lErro = CF("Ccl_Formata", CclInicial.Text, sCcl_I, iCclPreenchida_I)
    If lErro <> SUCESSO Then Error 13428

    lErro = CF("Ccl_Formata", CclFinal.Text, sCcl_F, iCclPreenchida_F)
    If lErro <> SUCESSO Then Error 13429

    If (iCclPreenchida_I = CCL_PREENCHIDA) And (iCclPreenchida_F = CCL_PREENCHIDA) Then
    
        If sCcl_I > sCcl_F Then Error 13430
    
    End If
    
    'grava os parâmetros no arquivo C
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 13431
    
    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then Error 13432

    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then Error 13433
    
    lErro = objRelOpcoes.IncluirParametro("NPERIODOINICIAL", CStr(PeriodoInicial.ItemData(PeriodoInicial.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 47165

    lErro = objRelOpcoes.IncluirParametro("NPERIODO", CStr(PeriodoFinal.ItemData(PeriodoFinal.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 47166

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13435

    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 19398
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", PeriodoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47168
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX3", PeriodoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47167

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCcl_I, iCclPreenchida_I, sCcl_F, iCclPreenchida_F)
    If lErro <> SUCESSO Then Error 13436

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 13426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 13427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_VAZIO", Err)
            PeriodoInicial.SetFocus

        Case 13428, 13429

        Case 13430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", Err)

        Case 13431, 13432, 13433, 13435, 13436, 19398, 47165, 47166, 47167, 47168
        
        Case 47169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_FINAL_VAZIO", Err)
            PeriodoFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168211)

    End Select

    Exit Function

End Function

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
'    If lErro <> SUCESSO Then Error 13437
'
'    'para cada centro de custo encontrado no bd
'    For Each objCcl In colCcl
'
'        sCclMascarado = String(STRING_CCL, 0)
'
'        'coloca a mascara no centro de custo
'        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
'        If lErro <> SUCESSO Then Error 13438
'
'        sCcl = "C" & objCcl.sCcl
'
'        sCclPai = String(STRING_CCL, 0)
'
'        'retorna o centro de custo/lucro "pai" da centro de custo/lucro em questão, se houver
'        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
'        If lErro <> SUCESSO Then Error 54703
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
'        Case 54703
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)
'
'        Case 13437
'
'        Case 13438
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168212)
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
''    If lErro <> SUCESSO Then Error 13437
''
''    For Each objCcl In colCcl
''
''        sCclMascarado = String(STRING_CCL, 0)
''
''        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
''        If lErro <> SUCESSO Then Error 13438
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
''        Case 13437
''
''        Case 13438
''            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168213)
''
''    End Select
''
''    Exit Function
''
''End Function

Function MostraExercicioPeriodos(iExercicio As Integer, iPer_I As Integer, iPer_F As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama PreencheComboPeriodos

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
    If lErro <> SUCESSO Then Error 13440

    MostraExercicioPeriodos = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodos:

    MostraExercicioPeriodos = Err

    Select Case Err

        Case 13440

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168214)

    End Select

    Exit Function

End Function


Function PreencheComboPeriodos(iExercicio As Integer, iPer_I As Integer, iPer_F As Integer) As Long
'lê os períodos do exercício 'iExercicio' preenchendo o combo de período
'seleciona o período 'iPeriodo'

Dim lErro As Long, iIndice As Integer
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheComboPeriodos

    PeriodoInicial.Clear
    PeriodoFinal.Clear

    'inicializar os periodos do exercicio selecionado no combo de exercícios
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 13441

    For iIndice = 1 To colPeriodos.Count
        
        Set objPeriodo = colPeriodos.Item(iIndice)
        
        PeriodoInicial.AddItem objPeriodo.sNomeExterno
        PeriodoInicial.ItemData(PeriodoInicial.NewIndex) = objPeriodo.iPeriodo
        
        PeriodoFinal.AddItem objPeriodo.sNomeExterno
        PeriodoFinal.ItemData(PeriodoFinal.NewIndex) = objPeriodo.iPeriodo
    
    Next

    'mostra o período inicial
    For iIndice = 0 To PeriodoInicial.ListCount - 1
        If PeriodoInicial.ItemData(iIndice) = iPer_I Then
            PeriodoInicial.ListIndex = iIndice
            Exit For
        End If
    Next

    'mostra o período final
    For iIndice = 0 To PeriodoFinal.ListCount - 1
        If PeriodoFinal.ItemData(iIndice) = iPer_F Then
            PeriodoFinal.ListIndex = iIndice
            Exit For
        End If
    Next
    
    PreencheComboPeriodos = SUCESSO

    Exit Function

Erro_PreencheComboPeriodos:

    PreencheComboPeriodos = Err

    Select Case Err

        Case 13441

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168215)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29591
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    Caption = gobjRelatorio.sCodRel
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 13452
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 13452
        
        Case 29591
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168216)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 13442

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 13443

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47051
        
        DescCclInic.Caption = ""
        DescCclFim.Caption = ""
    
    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 13442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 13443, 47051

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168217)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13444

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 13444

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168218)

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
    If ComboOpcoes.Text = "" Then Error 13425

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13445

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 13446

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47049
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 13425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 13445

        Case 13446, 47049
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168219)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47048
    
    DescCclInic.Caption = ""
    DescCclFim.Caption = ""
    ComboOpcoes.Text = ""
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47048
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168220)

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
    If lErro <> SUCESSO Then Error 54897
    
    'coloca a mascara na tela.
    CclInicial.Mask = sMascaraCcl
    
    Inicializa_Mascara_CclInicial = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_CclInicial:

    Inicializa_Mascara_CclInicial = Err
    
    Select Case Err
    
        Case 54897
            lErro = Rotina_Erro(vbOKOnly, "Erro_MascaraCcl", Err)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168221)
        
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
    If lErro <> SUCESSO Then Error 54898
    
    'coloca a mascara na tela.
    CclFinal.Mask = sMascaraCcl
    
    Inicializa_Mascara_CclFinal = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_CclFinal:

    Inicializa_Mascara_CclFinal = Err
    
    Select Case Err
    
        Case 54898
            lErro = Rotina_Erro(vbOKOnly, "Erro_MascaraCcl", Err)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168222)
        
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
        If lErro <> SUCESSO And lErro <> 87164 Then gError 87168
    
        'se o centro de custo/lucro não estiver cadastrado
        If lErro = 87164 Then gError 87169

        lErro = Ccl_Perde_Foco(CclFinal, DescCclFim, objCcl)
        If lErro <> SUCESSO Then gError 81170

    End If
    
    Exit Sub
    
Erro_CclFinal_Validate:

    Cancel = True


    Select Case gErr
        
        Case 87169
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CclFinal.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatado
                
                Call Chama_Tela("CclTela", objCcl)
                
            
            Else
                
            End If

        Case 87168, 81170
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168223)
        
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
        If lErro <> SUCESSO And lErro <> 87164 Then gError 87166
    
        'se o centro de custo/lucro não estiver cadastrado
        If lErro = 87164 Then gError 87167

        lErro = Ccl_Perde_Foco(CclInicial, DescCclInic, objCcl)
        If lErro <> SUCESSO Then gError 87171

    End If
        
    Exit Sub
    
Erro_CclInicial_Validate:

    Cancel = True


    Select Case gErr
            
        Case 87167
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", CclInicial.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatado
                
                Call Chama_Tela("CclTela", objCcl)
                        
            End If

        Case 87166, 87171
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168224)
        
    End Select

    Exit Sub
    
End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click

    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    If giCarregando = CANCELA Then
    
        lErro = PreencheComboPeriodos(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1, 1)
        If lErro <> SUCESSO Then Error 13449
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 13449

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168225)

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

Dim lErro As Long, iIndice As Integer
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection

On Error GoTo Erro_Form_Load

    giCarregando = CANCELA
    giFocoInicial = 1

    'inicializa a mascara de centro de custo/lucro inicial
    lErro = Inicializa_Mascara_CclInicial()
    If lErro <> SUCESSO Then Error 54893
    
    'inicializa a mascara de centro de custo/lucro final
    lErro = Inicializa_Mascara_CclFinal()
    If lErro <> SUCESSO Then Error 54894

'    'Inicializa a Lista de Centros de Custo
'    lErro = CF("Carga_Arvore_Ccl", TvwCcls.Nodes)
'    If lErro <> SUCESSO Then Error 13453

    Set objEventoCclDe = New AdmEvento
    Set objEventoCclAte = New AdmEvento

    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 13454
    
    For iIndice = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iIndice)
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
 
        Case 13453, 13454, 54893, 54894

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168226)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim lErro As Long
Dim sCcl As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_TvwCcls_NodeClick
    
    objCcl.sCcl = Right(Node.Key, Len(Node.Key) - 1)
    
    If giFocoInicial = 1 Then
        lErro = Ccl_Perde_Foco(CclInicial, DescCclInic, objCcl)
        If lErro <> SUCESSO Then gError 87160
    
    Else
        lErro = Ccl_Perde_Foco(CclFinal, DescCclFim, objCcl)
        If lErro <> SUCESSO Then gError 87161
    
    End If
    
    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case gErr

        Case 87160, 87161

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168227)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DESP_PER_CCL
    Set Form_Load_Ocx = Me
    Caption = "Balancete por Centro de Custo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDespPerCcl"
    
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



'Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label7, Source, X, Y)
'End Sub
'
'Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label8, Source, X, Y)
'End Sub
'
'Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168228)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168229)
        
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
