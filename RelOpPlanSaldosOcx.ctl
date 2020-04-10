VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpPlanSaldosOcx 
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   8310
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPlanSaldosOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPlanSaldosOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPlanSaldosOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPlanSaldosOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpPlanSaldosOcx.ctx":0994
      Left            =   1050
      List            =   "RelOpPlanSaldosOcx.ctx":0996
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   930
      Width           =   1545
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPlanSaldosOcx.ctx":0998
      Left            =   1050
      List            =   "RelOpPlanSaldosOcx.ctx":099A
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2535
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
      Height          =   585
      Left            =   4005
      Picture         =   "RelOpPlanSaldosOcx.ctx":099C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contas"
      Height          =   2535
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   7995
      Begin VB.CheckBox CheckZeradas 
         Caption         =   "Exibir as contas zeradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1215
         Width           =   3615
      End
      Begin VB.CheckBox CheckSintInt 
         Caption         =   "Exibir as contas sintéticas intermediárias"
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
         TabIndex        =   7
         Top             =   1695
         Width           =   3840
      End
      Begin VB.TextBox NivelMaximo 
         Height          =   285
         Left            =   2190
         MaxLength       =   1
         TabIndex        =   8
         Top             =   2115
         Width           =   255
      End
      Begin MSMask.MaskEdBox ContaInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaFinal 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescCtaFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2715
         TabIndex        =   27
         Top             =   840
         Width           =   4995
      End
      Begin VB.Label DescCtaInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2715
         TabIndex        =   26
         Top             =   360
         Width           =   4995
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
         Left            =   90
         TabIndex        =   25
         Top             =   360
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
         Left            =   195
         TabIndex        =   24
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Nível Máximo de Conta:"
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
         Left            =   105
         TabIndex        =   23
         Top             =   2130
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo do Relatório"
      Height          =   675
      Left            =   120
      TabIndex        =   21
      Top             =   4035
      Width           =   7995
      Begin VB.OptionButton OptionMovimentacao 
         Caption         =   "Movimentações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2535
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptionSaldos 
         Caption         =   "Saldos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   870
         TabIndex        =   9
         Top             =   255
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Periodos"
      Height          =   645
      Left            =   2940
      TabIndex        =   18
      Top             =   735
      Width           =   5175
      Begin VB.ComboBox PeriodoFinal 
         Height          =   315
         Left            =   3345
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   255
         Width           =   1695
      End
      Begin VB.ComboBox PeriodoInicial 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Left            =   2805
         TabIndex        =   20
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label7 
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
         Left            =   180
         TabIndex        =   19
         Top             =   270
         Width           =   585
      End
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   3030
      Left            =   5700
      TabIndex        =   11
      Top             =   1665
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5345
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
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
      Left            =   150
      TabIndex        =   30
      Top             =   975
      Width           =   855
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   29
      Top             =   255
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
      Left            =   5685
      TabIndex        =   28
      Top             =   1455
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "RelOpPlanSaldosOcx"
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

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCheck As String, sCta_I As String, sCta_F As String

On Error GoTo Erro_PreencherRelOp

    sCheck = String(1, 0)
    sCta_I = String(STRING_CONTA, 0)
    sCta_F = String(STRING_CONTA, 0)

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 13239

    'período inicial não pode ser vazio
    If PeriodoInicial.Text = "" Then Error 13240
    
    'periodo final não pode estar vazio
    If PeriodoFinal.Text = "" Then Error 19403
    
   If PeriodoInicial.ItemData(PeriodoInicial.ListIndex) > PeriodoFinal.ItemData(PeriodoFinal.ListIndex) Then Error 28906
    
    lErro = Formata_E_Critica_Contas(sCta_I, sCta_F)
    If lErro <> SUCESSO Then Error 13205

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 13206

    'lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(giFilialEmpresa))
    'If lErro <> AD_BOOL_TRUE Then Error 7219

    'lErro = objRelOpcoes.IncluirParametro("TNOMEFILIAL", CStr(gsNomeFilialEmpresa))
    'If lErro <> AD_BOOL_TRUE Then Error 7220

    'transforma Saldos em "S" e Movimentação em "M" e define qual o .tsk a ser executado
    If OptionSaldos.Value = True Then
        sCheck = "S"
        gobjRelatorio.sNomeTsk = "salmov"
    Else
        sCheck = "M"
        gobjRelatorio.sNomeTsk = "salmov"
    End If

    lErro = objRelOpcoes.IncluirParametro("TSALDOSMOV", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 13207

    lErro = objRelOpcoes.IncluirParametro("TCTAINIC", sCta_I)
    If lErro <> AD_BOOL_TRUE Then Error 13208

    lErro = objRelOpcoes.IncluirParametro("TCTAFIM", sCta_F)
    If lErro <> AD_BOOL_TRUE Then Error 13209

    'transforma o valor do check box CheckZeradas em "S" ou "N"
    If CheckZeradas.Value = 0 Then
        sCheck = "N"
    Else
        sCheck = "S"
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TZERADAS", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 13210
    
    'transforma o valor do check box CheckSintInt em "S" ou "N"
    If CheckSintInt.Value = 0 Then
        sCheck = "N"
    Else
        sCheck = "S"
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TSINTINT", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 13211

    lErro = objRelOpcoes.IncluirParametro("NCTANIVMAX", NivelMaximo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13212

    lErro = objRelOpcoes.IncluirParametro("NPERIODO", CStr(PeriodoFinal.ItemData(PeriodoFinal.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13213
    
    lErro = objRelOpcoes.IncluirParametro("NPERIODOINICIAL", CStr(PeriodoInicial.ItemData(PeriodoInicial.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 19404

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13214
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 19402
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", PeriodoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54705
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX3", PeriodoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54706

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCta_I, sCta_F)
    If lErro <> SUCESSO Then Error 13215

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err
    
    Select Case Err

        Case 13239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 13241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_VAZIO", Err)
            PeriodoInicial.SetFocus
                
        Case 13205

        Case 13206

        Case 13207, 13208, 13209, 13210, 13211, 13212, 13213, 13214

        Case 13215, 19402, 19404, 54705, 54706
                
        Case 19403
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_FINAL_VAZIO", Err)
            PeriodoFinal.SetFocus
        
        Case 28906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_INICIAL_MAIOR", Err)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171184)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long, iExercicio As Integer
Dim iPer_I As Integer
Dim iPer_F As Integer
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 13216

    'pega a informação se é Saldo ou Movimentação
    lErro = objRelOpcoes.ObterParametro("TSALDOSMOV", sParam)

    If sParam = "S" Then
        OptionSaldos.Value = True
    Else
        OptionMovimentacao.Value = True
    End If

    'pega Conta Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAINIC", sParam)
    If lErro <> SUCESSO Then Error 13217

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then Error 13218

    'pega Conta Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAFIM", sParam)
    If lErro <> SUCESSO Then Error 13219

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaFinal, DescCtaFim)
    If lErro <> SUCESSO Then Error 13220

    'imprimir contas zeradas
    lErro = objRelOpcoes.ObterParametro("TZERADAS", sParam)
    If lErro <> SUCESSO Then Error 13221

    If sParam = "S" Then CheckZeradas.Value = 1
    If sParam = "N" Then CheckZeradas.Value = 0
    
    'imprimir contas sintéticas intermediárias
    lErro = objRelOpcoes.ObterParametro("TSINTINT", sParam)
    If lErro <> SUCESSO Then Error 13222

    If sParam = "S" Then CheckSintInt.Value = 1
    If sParam = "N" Then CheckSintInt.Value = 0
    
    'limitar nível máximo de conta
    lErro = objRelOpcoes.ObterParametro("NCTANIVMAX", sParam)
    If lErro <> SUCESSO Then Error 13223

    NivelMaximo.Text = sParam

    'período final
    lErro = objRelOpcoes.ObterParametro("NPERIODO", sParam)
    If lErro <> SUCESSO Then Error 13224

    iPer_F = CInt(sParam)

    'período inicial
    lErro = objRelOpcoes.ObterParametro("NPERIODOINICIAL", sParam)
    If lErro <> SUCESSO Then Error 19405

    iPer_I = CInt(sParam)

    
    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 13225

    iExercicio = CInt(sParam)
        
    lErro = MostraExercicioPeriodo(iExercicio, iPer_I, iPer_F)
    If lErro <> SUCESSO Then Error 13226

    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err
    
    Select Case Err

        Case 13216

        Case 13217, 13219, 13221, 13222, 13223, 13224, 13225

        Case 13218, 13220

        Case 13226, 19405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171185)

    End Select
    
    Exit Function
    
End Function

Private Function Formata_E_Critica_Contas(sCta_I As String, sCta_F As String) As Long
'retorna em sCta_I e sCta_F as contas (inicial e final) formatadas
'Verifica se a conta inicial é maior que a conta final

Dim iCtaPreenchida_I As Integer, iCtaPreenchida_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Contas

    'formata a Conta Inicial
    lErro = CF("Conta_Formata", ContaInicial.Text, sCta_I, iCtaPreenchida_I)
    If lErro <> SUCESSO Then Error 13227
    If iCtaPreenchida_I <> CONTA_PREENCHIDA Then sCta_I = ""

    'formata a Conta Final
    lErro = CF("Conta_Formata", ContaFinal.Text, sCta_F, iCtaPreenchida_F)
    If lErro <> SUCESSO Then Error 13228
    If iCtaPreenchida_F <> CONTA_PREENCHIDA Then sCta_F = ""

    'se ambas as contas estão preenchidas, a conta inicial não pode ser maior que a final
    If iCtaPreenchida_I = CONTA_PREENCHIDA And iCtaPreenchida_F = CONTA_PREENCHIDA Then

        If sCta_I > sCta_F Then Error 13229

    End If

    Formata_E_Critica_Contas = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Contas:

    Formata_E_Critica_Contas = Err

    Select Case Err

        Case 13227
            ContaInicial.SetFocus

        Case 13228
            ContaFinal.SetFocus

        Case 13229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171186)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCta_I As String, sCta_F As String) As Long
'monta a expressão de seleção
'recebe as contas inicial e final no formato do BD

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sCta_I <> "" Then sExpressao = "Conta >= " & Forprint_ConvTexto(sCta_I)

    If sCta_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Conta <= " & Forprint_ConvTexto(sCta_F)

    End If

    If CheckZeradas.Value = 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CtaNaoZeradaPer"

    End If

    If CheckSintInt.Value = 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "SemCtaSintInt"

    End If

    If Len(NivelMaximo.Text) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NivelConta <= " & Forprint_ConvInt(CInt(NivelMaximo.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171187)

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
        If lErro <> SUCESSO Then Error 13230

    Else

        lErro = CF("Traz_Conta_MaskEd", sConta, ContaFinal, DescCtaFim)
        If lErro <> SUCESSO Then Error 13231

    End If

    Traz_Conta_Tela = SUCESSO

    Exit Function

Erro_Traz_Conta_Tela:

    Traz_Conta_Tela = Err

    Select Case Err

        Case 13230

        Case 13231

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171188)

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

    If Not (gobjRelatorio Is Nothing) Then Error 24977
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 13250
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 13250
        
        Case 24977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171189)

    End Select

    Exit Function

End Function

Function MostraExercicioPeriodo(iExercicio As Integer, iPer_I As Integer, iPer_F As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama Preenche combos do periodos

Dim iConta As Integer, lErro As Long

On Error GoTo Erro_MostraExercicioPeriodo

    giCarregando = OK

    For iConta = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iConta) = iExercicio Then
            ComboExercicio.ListIndex = iConta
            Exit For
        End If
    Next

    lErro = PreencheComboPeriodos(iExercicio, iPer_I, iPer_F)
    If lErro <> SUCESSO Then Error 13233

    MostraExercicioPeriodo = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodo:

    MostraExercicioPeriodo = Err

    Select Case Err

        Case 13233

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171190)

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

    PeriodoInicial.Clear
    PeriodoFinal.Clear

    'inicializar os periodos do exercicio selecionado no combo de exercícios
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 13234

    'adiciona os periodos na combo e posiciona com o item passado nos parametros iPer_I e iPer_F
    For Each objPeriodo In colPeriodos
    
        PeriodoInicial.AddItem objPeriodo.sNomeExterno
        PeriodoInicial.ItemData(PeriodoInicial.NewIndex) = objPeriodo.iPeriodo
        
        If PeriodoInicial.ItemData(PeriodoInicial.NewIndex) = iPer_I Then
            PeriodoInicial.ListIndex = PeriodoInicial.NewIndex
        End If
        
        PeriodoFinal.AddItem objPeriodo.sNomeExterno
        PeriodoFinal.ItemData(PeriodoFinal.NewIndex) = objPeriodo.iPeriodo
        
        If PeriodoFinal.ItemData(PeriodoFinal.NewIndex) = iPer_F Then
            PeriodoFinal.ListIndex = PeriodoFinal.NewIndex
        End If
    
    Next
    
    PreencheComboPeriodos = SUCESSO

    Exit Function

Erro_PreencheComboPeriodos:

    PreencheComboPeriodos = Err

    Select Case Err

        Case 13234

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171191)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 13235

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 13236

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47090
    
        DescCtaInic.Caption = ""
        DescCtaFim.Caption = ""
        CheckZeradas.Value = 0
        CheckSintInt.Value = 0
        OptionSaldos.Value = True

    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 13235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 13236

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171192)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13237

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 13237

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171193)

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
    If ComboOpcoes.Text = "" Then Error 13238

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13241

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 13242

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47088
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 13238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 13241, 13242

        Case 13243, 47088
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171194)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47081
    
    DescCtaInic.Caption = ""
    DescCtaFim.Caption = ""
    CheckZeradas.Value = 0
    CheckSintInt.Value = 0
    OptionSaldos.Value = True
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47081
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171195)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click

    If ComboExercicio.ListIndex = -1 Then Exit Sub

    If giCarregando = CANCELA Then

        lErro = PreencheComboPeriodos(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1, 1)
        If lErro <> SUCESSO Then Error 13243
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 13243

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171196)

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
    If lErro <> SUCESSO Then Error 13246

    Exit Sub

Erro_ContaFinal_Validate:

    Cancel = True


    Select Case Err

        Case 13246

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171197)

    End Select

    Exit Sub

End Sub

Private Sub ContaInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ContaInicial_Validate

    giFocoInicial = 1

    lErro = CF("Conta_Perde_Foco", ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then Error 13247

    Exit Sub

Erro_ContaInicial_Validate:

    Cancel = True


    Select Case Err

        Case 13247

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171198)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long, iConta As Integer
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection

On Error GoTo Erro_Form_Load

    OptionSaldos.Value = True

    giCarregando = CANCELA
    giFocoInicial = 1
      
    'inicializa a mascara de conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaInicial)
    If lErro <> SUCESSO Then Error 13248

    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaFinal)
    If lErro <> SUCESSO Then Error 13249

'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 13251

    Set objEventoContaDe = New AdmEvento
    Set objEventoContaAte = New AdmEvento

    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 13252
    
    For iConta = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iConta)
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next
       
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 13248

        Case 13249
        
        Case 13251
        
        Case 13252
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171199)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub NivelMaximo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NivelMaximo_Validate

    'lote inicial deve estar entre 1 e 9999
    If NivelMaximo.Text <> "" Then
        lErro = Valor_Critica(NivelMaximo.Text)
        If lErro <> SUCESSO Then Error 54885
    End If

    Exit Sub

Erro_NivelMaximo_Validate:

    Cancel = True


    Select Case Err
        
        Case 54885
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171200)

    End Select

    Exit Sub
    
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 40824
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 40824
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171201)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_TvwContas_NodeClick

    sConta = Right(Node.Key, Len(Node.Key) - 1)

    lErro = Traz_Conta_Tela(sConta)
    If lErro <> SUCESSO Then Error 13254

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 13254

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171202)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PLAN_SALDOS
    Set Form_Load_Ocx = Me
    Caption = "Saldos e Movimentações do Exercício"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPlanSaldos"
    
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
'
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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
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


