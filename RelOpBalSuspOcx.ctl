VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpBalSuspOcx 
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   8100
   Begin VB.Frame Frame2 
      Caption         =   "Diário Geral"
      Height          =   795
      Left            =   90
      TabIndex        =   24
      Top             =   4005
      Width           =   7920
      Begin VB.TextBox PrimeiraFolha 
         Height          =   285
         Left            =   6210
         TabIndex        =   7
         Top             =   300
         Width           =   510
      End
      Begin VB.TextBox Diario 
         Height          =   285
         Left            =   2595
         TabIndex        =   6
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label8 
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
         Left            =   3930
         TabIndex        =   26
         Top             =   330
         Width           =   2250
      End
      Begin VB.Label Label7 
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
         Left            =   1035
         TabIndex        =   25
         Top             =   345
         Width           =   1545
      End
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpBalSuspOcx.ctx":0000
      Left            =   1020
      List            =   "RelOpBalSuspOcx.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   855
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contas"
      Height          =   2535
      Left            =   90
      TabIndex        =   14
      Top             =   1320
      Width           =   7920
      Begin VB.TextBox NivelMaximo 
         Height          =   285
         Left            =   2220
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2100
         Width           =   255
      End
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
         TabIndex        =   3
         Top             =   1200
         Width           =   4800
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
         TabIndex        =   4
         Top             =   1680
         Width           =   4905
      End
      Begin MSMask.MaskEdBox ContaFinal 
         Height          =   315
         Left            =   945
         TabIndex        =   27
         Top             =   765
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaInicial 
         Height          =   315
         Left            =   930
         TabIndex        =   28
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   19
         Top             =   2130
         Width           =   2055
      End
      Begin VB.Label DescCtaFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2955
         TabIndex        =   18
         Top             =   765
         Width           =   4845
      End
      Begin VB.Label DescCtaInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2955
         TabIndex        =   17
         Top             =   270
         Width           =   4845
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   300
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
         Left            =   390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   795
         Width           =   615
      End
   End
   Begin VB.ComboBox ComboPeriodo 
      Height          =   315
      Left            =   3690
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   900
      Width           =   1695
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
      Left            =   4110
      Picture         =   "RelOpBalSuspOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   90
      Width           =   1500
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpBalSuspOcx.ctx":0106
      Left            =   1020
      List            =   "RelOpBalSuspOcx.ctx":0108
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   315
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5850
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpBalSuspOcx.ctx":010A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1080
         Picture         =   "RelOpBalSuspOcx.ctx":0288
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpBalSuspOcx.ctx":07BA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpBalSuspOcx.ctx":0944
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   2745
      Left            =   5580
      TabIndex        =   29
      Top             =   1110
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   4842
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
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
      Left            =   5610
      TabIndex        =   23
      Top             =   885
      Visible         =   0   'False
      Width           =   1410
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
      TabIndex        =   22
      Top             =   930
      Width           =   855
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
      Left            =   2850
      TabIndex        =   21
      Top             =   930
      Width           =   735
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
      TabIndex        =   20
      Top             =   345
      Width           =   615
   End
End
Attribute VB_Name = "RelOpBalSuspOcx"
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

Function MostraExercicioPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama PreencheComboPeriodo

Dim iConta As Integer, lErro As Long

On Error GoTo Erro_MostraExercicioPeriodo
    
    giCarregando = OK

    For iConta = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iConta) = iExercicio Then
            ComboExercicio.ListIndex = iConta
            Exit For
        End If
    Next

    lErro = PreencheComboPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then gError 80349

    MostraExercicioPeriodo = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodo:

    MostraExercicioPeriodo = gErr

    Select Case gErr

        Case 80349

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167287)

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
    If lErro <> SUCESSO Then gError 80350

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

    PreencheComboPeriodo = gErr

    Select Case gErr

        Case 80350

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167288)

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
        If lErro <> SUCESSO Then gError 80351

    Else

        lErro = CF("Traz_Conta_MaskEd", sConta, ContaFinal, DescCtaFim)
        If lErro <> SUCESSO Then gError 80352

    End If

    Traz_Conta_Tela = SUCESSO

    Exit Function

Erro_Traz_Conta_Tela:

    Traz_Conta_Tela = gErr

    Select Case gErr

        Case 80351

        Case 80352

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167289)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 80353
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 80354
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 80353
        
        Case 80354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167290)

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
    If lErro <> SUCESSO Then gError 80355
    If iCtaPreenchida_I <> CONTA_PREENCHIDA Then sCta_I = ""

    'formata a Conta Final
    lErro = CF("Conta_Formata", ContaFinal.Text, sCta_F, iCtaPreenchida_F)
    If lErro <> SUCESSO Then gError 80356
    If iCtaPreenchida_F <> CONTA_PREENCHIDA Then sCta_F = ""

    'se ambas as contas estão preenchidas, a conta inicial não pode ser maior que a final
    If iCtaPreenchida_I = CONTA_PREENCHIDA And iCtaPreenchida_F = CONTA_PREENCHIDA Then

        If sCta_I > sCta_F Then gError 80357

    End If

    Formata_E_Critica_Contas = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Contas:

    Formata_E_Critica_Contas = gErr

    Select Case gErr

        Case 80355
            ContaInicial.SetFocus

        Case 80356
            ContaFinal.SetFocus

        Case 80357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIAL_MAIOR", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167291)

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

'    If CheckZeradas.Value = 0 Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "CtaNaoZeradaPer"
'
'    End If

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

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167292)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCheck As String, sCta_I As String, sCta_F As String

On Error GoTo Erro_PreencherRelOp

    sCheck = String(1, 0)
    sCta_I = String(STRING_CONTA, 0)
    sCta_F = String(STRING_CONTA, 0)

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then gError 80357

    'período não pode ser vazio
    If ComboPeriodo.Text = "" Then gError 80358

    lErro = Formata_E_Critica_Contas(sCta_I, sCta_F)
    If lErro <> SUCESSO Then gError 80359

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 80360

    'lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(giFilialEmpresa))
    'If lErro <> AD_BOOL_TRUE Then Error 7238

    'lErro = objRelOpcoes.IncluirParametro("TNOMEFILIAL", CStr(gsNomeFilialEmpresa))
    'If lErro <> AD_BOOL_TRUE Then Error 7239

    lErro = objRelOpcoes.IncluirParametro("TCTAINIC", sCta_I)
    If lErro <> AD_BOOL_TRUE Then gError 80361

    lErro = objRelOpcoes.IncluirParametro("TCTAFIM", sCta_F)
    If lErro <> AD_BOOL_TRUE Then gError 80362

    'transforma o valor do check box CheckZeradas em "S" ou "N"
    If CheckZeradas.Value = 0 Then
        sCheck = "N"
    Else
        sCheck = "S"
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TZERADAS", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 80363
    
    'transforma o valor do check box CheckSintInt em "S" ou "N"
    If CheckSintInt.Value = 0 Then
        sCheck = "N"
    Else
        sCheck = "S"
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TSINTINT", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 80364

    lErro = objRelOpcoes.IncluirParametro("NCTANIVMAX", NivelMaximo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 80365

    lErro = objRelOpcoes.IncluirParametro("NPERIODO", CStr(ComboPeriodo.ItemData(ComboPeriodo.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 80366

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 80367
 
    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then gError 80368
    
    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", ComboPeriodo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 80369

    'Preencher RelOpcoes
    lErro = objRelOpcoes.IncluirParametro("NPAGRELINI", PrimeiraFolha.Text)
    If lErro <> AD_BOOL_TRUE Then gError 80371

    lErro = objRelOpcoes.IncluirParametro("NNUMDIARIO", Diario.Text)
    If lErro <> AD_BOOL_TRUE Then gError 80372
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCta_I, sCta_F)
    If lErro <> SUCESSO Then gError 80370
      
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr

        Case 80357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 80358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)
            ComboPeriodo.SetFocus

        Case 80359

        Case 80360

        Case 80361, 80362, 80363, 80364, 80365, 80366, 80367

        Case 80368, 80369, 80370, 80371, 80372, 80373

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167293)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long, iExercicio As Integer, iPeriodo As Integer
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 80374

    'pega Conta Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAINIC", sParam)
    If lErro <> SUCESSO Then gError 80375

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then gError 80376

    'pega Conta Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAFIM", sParam)
    If lErro <> SUCESSO Then gError 80377

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaFinal, DescCtaFim)
    If lErro <> SUCESSO Then gError 80378

    'imprimir contas zeradas
    lErro = objRelOpcoes.ObterParametro("TZERADAS", sParam)
    If lErro <> SUCESSO Then gError 80379

    If sParam = "S" Then CheckZeradas.Value = 1

    'imprimir contas sintéticas intermediárias
    lErro = objRelOpcoes.ObterParametro("TSINTINT", sParam)
    If lErro <> SUCESSO Then gError 80380

    If sParam = "S" Then CheckSintInt.Value = 1

    'limitar nível máximo de conta
    lErro = objRelOpcoes.ObterParametro("NCTANIVMAX", sParam)
    If lErro <> SUCESSO Then gError 80381

    NivelMaximo.Text = sParam

    'período
    lErro = objRelOpcoes.ObterParametro("NPERIODO", sParam)
    If lErro <> SUCESSO Then gError 80382

    iPeriodo = CInt(sParam)

    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then gError 80383

    iExercicio = CInt(sParam)

    lErro = MostraExercicioPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then gError 80384

    'pega primeira folha e exibe
    lErro = objRelOpcoes.ObterParametro("NPAGRELINI", sParam)
    If lErro <> SUCESSO Then gError 80385

    PrimeiraFolha.Text = sParam

    'pega número diário e exibe
    lErro = objRelOpcoes.ObterParametro("NNUMDIARIO", sParam)
    If lErro <> SUCESSO Then gError 80386

    Diario.Text = sParam

    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr
    
    Select Case gErr

        Case 80374

        Case 80375, 80376, 80377, 80378, 80379, 80380, 80381

        Case 80382, 80383

        Case 80384, 80385, 80386

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167294)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 80387

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 80388

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 80389
    
        DescCtaInic.Caption = ""
        DescCtaFim.Caption = ""
        CheckZeradas.Value = 0
        CheckSintInt.Value = 0

    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case gErr

        Case 80387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 80388, 80389

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167295)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 80390

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 80390

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167296)

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
    If ComboOpcoes.Text = "" Then gError 80391

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 80392

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 80393

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 80394
        
    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 80391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 80392

        Case 80393, 80394
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167297)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 80395
    
    DescCtaInic.Caption = ""
    DescCtaFim.Caption = ""
    CheckZeradas.Value = 0
    CheckSintInt.Value = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 80395
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167298)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click

    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    If giCarregando = CANCELA Then

        lErro = PreencheComboPeriodo(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1)
        If lErro <> SUCESSO Then gError 80396
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

    giCarregando = CANCELA
    
Erro_ComboExercicio_Click:

    Select Case gErr

        Case 80396

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167299)

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
    If lErro <> SUCESSO Then gError 80397

    Exit Sub

Erro_ContaFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 80397

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167300)

    End Select

    Exit Sub

End Sub

Private Sub ContaInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ContaInicial_Validate

    giFocoInicial = 1

    lErro = CF("Conta_Perde_Foco", ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then gError 80398

    Exit Sub

Erro_ContaInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 80398

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167301)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long, iConta As Integer
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection

On Error GoTo Erro_RelOpBalVerif_Form_Load

    giCarregando = CANCELA
    giFocoInicial = 1

    'inicializa a mascara de conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaInicial)
    If lErro <> SUCESSO Then gError 80399

    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaFinal)
    If lErro <> SUCESSO Then gError 80400

'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then gError 80401

    Set objEventoContaDe = New AdmEvento
    Set objEventoContaAte = New AdmEvento

    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then gError 80402
    
    For Each objExercicio In colExerciciosAbertos
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next
    
    ComboExercicio.ListIndex = -1
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_RelOpBalVerif_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 80399

        Case 80400
        
        Case 80401

        Case 80402

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167302)

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
        If lErro <> SUCESSO Then gError 80403
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case gErr
    
        Case 80403
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167303)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_TvwContas_NodeClick

    sConta = Right(Node.Key, Len(Node.Key) - 1)

    lErro = Traz_Conta_Tela(sConta)
    If lErro <> SUCESSO Then gError 80405

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case gErr

        Case 80405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167304)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoContaDe = Nothing
    Set objEventoContaAte = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_BALANCO_VERIFIC
    Set Form_Load_Ocx = Me
    Caption = "Balancete de Suspensão"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBalSusp"
    
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




'Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label4, Source, X, Y)
'End Sub
'
'Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub

Private Sub DescCtaInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCtaInic, Source, X, Y)
End Sub

Private Sub DescCtaInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCtaInic, Button, Shift, X, Y)
End Sub

Private Sub DescCtaFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCtaFim, Source, X, Y)
End Sub

Private Sub DescCtaFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCtaFim, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
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





