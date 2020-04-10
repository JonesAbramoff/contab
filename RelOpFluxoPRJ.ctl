VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpFluxoPRJ 
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   6285
   Begin VB.Frame Frame1 
      Caption         =   "Natureza"
      Height          =   1140
      Left            =   120
      TabIndex        =   21
      Top             =   2955
      Width           =   6075
      Begin MSMask.MaskEdBox NaturezaDe 
         Height          =   315
         Left            =   480
         TabIndex        =   7
         Top             =   300
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NaturezaAte 
         Height          =   315
         Left            =   480
         TabIndex        =   8
         Top             =   690
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelNaturezaAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   75
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   735
         Width           =   345
      End
      Begin VB.Label LabelNaturezaDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelNaturezaAteDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2010
         TabIndex        =   23
         Top             =   690
         Width           =   3990
      End
      Begin VB.Label LabelNaturezaDeDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2010
         TabIndex        =   22
         Top             =   300
         Width           =   3990
      End
   End
   Begin VB.ComboBox Etapa 
      Height          =   315
      Left            =   3615
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2475
      Width           =   2550
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   1395
      Left            =   105
      TabIndex        =   16
      Top             =   825
      Width           =   3615
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   720
         TabIndex        =   1
         Top             =   285
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataInicial 
         Height          =   300
         Left            =   1875
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   285
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   720
         TabIndex        =   3
         Top             =   810
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataFinal 
         Height          =   300
         Left            =   1845
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   810
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelDataDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   375
         TabIndex        =   18
         Top             =   330
         Width           =   285
      End
      Begin VB.Label LabelDataAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   300
         TabIndex        =   17
         Top             =   855
         Width           =   330
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpFluxoPRJ.ctx":0000
      Left            =   840
      List            =   "RelOpFluxoPRJ.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2916
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
      Left            =   4590
      Picture         =   "RelOpFluxoPRJ.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1050
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4020
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpFluxoPRJ.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpFluxoPRJ.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpFluxoPRJ.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RelOpFluxoPRJ.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Projeto 
      Height          =   300
      Left            =   810
      TabIndex        =   5
      Top             =   2475
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   529
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Etapa:"
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
      Index           =   41
      Left            =   2985
      TabIndex        =   20
      Top             =   2490
      Width           =   570
   End
   Begin VB.Label LabelProjeto 
      AutoSize        =   -1  'True
      Caption         =   "Projeto:"
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
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   19
      Top             =   2520
      Width           =   675
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
      Index           =   0
      Left            =   135
      TabIndex        =   15
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpFluxoPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim glNumIntPRJ As Long
Dim glNumIntPRJEtapa As Long

Private WithEvents objEventoNaturezaDe As AdmEvento
Attribute objEventoNaturezaDe.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaAte As AdmEvento
Attribute objEventoNaturezaAte.VB_VarHelpID = -1

                   
Dim sProjetoAnt As String
Dim sEtapaAnt As String

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoNaturezaDe = Nothing
    Set objEventoNaturezaAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 194267
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 194268
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 194267
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 194268
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194269)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção
'recebe os produtos inicial e final no formato do BD

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 194270)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sNatDe As String, sNatAte As String) As Long

Dim lErro As Long
Dim iNaturezaPreenchida As Integer
Dim sNaturezaFormatada As String

On Error GoTo Erro_Formata_E_Critica_Parametros

    If StrParaDate(DataInicial.Text) <> DATA_NULA And StrParaDate(DataFinal.Text) <> DATA_NULA Then
        If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 194271
    End If
    
    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, NaturezaDe.Text, sNatDe, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, NaturezaAte.Text, sNatAte, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Len(Trim(sNatDe)) <> 0 And Len(Trim(sNatAte)) <> 0 Then
        If sNatDe > sNatAte Then gError 216110
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 194271
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
            
        Case 216110 'ERRO_NATUREZA_INICIAL_MAIOR
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZA_INICIAL_MAIOR", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194272)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lNumIntRel As Long
Dim lNumIntDocProjeto As Long
Dim lNumIntDocEtapa As Long
Dim sNatDe As String, sNatAte As String

On Error GoTo Erro_PreencherRelOp
           
    lErro = Formata_E_Critica_Parametros(sNatDe, sNatAte)
    If lErro <> SUCESSO Then gError 194273

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 194274
    
    lErro = objRelOpcoes.IncluirParametro("DDATAINI", Format(StrParaDate(DataInicial.Text), "dd/mm/yyyy"))
    If lErro <> AD_BOOL_TRUE Then gError 194275
    
    lErro = objRelOpcoes.IncluirParametro("DDATAFIM", Format(StrParaDate(DataFinal.Text), "dd/mm/yyyy"))
    If lErro <> AD_BOOL_TRUE Then gError 194276

    lErro = objRelOpcoes.IncluirParametro("TPROJETO", Projeto.Text)
    If lErro <> AD_BOOL_TRUE Then gError 194277

    lErro = objRelOpcoes.IncluirParametro("TETAPA", Etapa.Text)
    If lErro <> AD_BOOL_TRUE Then gError 194278

    lErro = objRelOpcoes.IncluirParametro("NPROJETO", CStr(glNumIntPRJ))
    If lErro <> AD_BOOL_TRUE Then gError 194279

    lErro = objRelOpcoes.IncluirParametro("NETAPA", CStr(glNumIntPRJEtapa))
    If lErro <> AD_BOOL_TRUE Then gError 194280
    
    lErro = objRelOpcoes.IncluirParametro("TNATDE", sNatDe)
    If lErro <> AD_BOOL_TRUE Then gError 194278
    
    lErro = objRelOpcoes.IncluirParametro("TNATATE", sNatAte)
    If lErro <> AD_BOOL_TRUE Then gError 194278
    
    If bExecutando Then
    
        lErro = CF("RelPRJCustoReceita_Prepara", lNumIntRel, glNumIntPRJ, glNumIntPRJEtapa, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), sNatDe, sNatAte)
        If lErro <> SUCESSO Then gError 194281
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 194282
    
    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 194273 To 194282
            'erro tratado nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194283)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim objProjeto As New ClassProjetos
Dim objEtapa As New ClassPRJEtapas
Dim sNaturezaEnxuta As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 194284

    'pega a Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINI", sParam)
    If lErro <> SUCESSO Then gError 194285
    Call DateParaMasked(DataInicial, StrParaDate(sParam))
    
    'pega a Data Final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 194286
    Call DateParaMasked(DataFinal, StrParaDate(sParam))
    
    lErro = objRelOpcoes.ObterParametro("TNATDE", sParam)
    If lErro <> SUCESSO Then gError 194286
    sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, sParam, sNaturezaEnxuta)
    If lErro <> SUCESSO Then gError 207582
    NaturezaDe.PromptInclude = False
    NaturezaDe.Text = sNaturezaEnxuta
    NaturezaDe.PromptInclude = True
    Call NaturezaDe_Validate(bSGECancelDummy)
    
    lErro = objRelOpcoes.ObterParametro("TNATATE", sParam)
    If lErro <> SUCESSO Then gError 194286
    sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, sParam, sNaturezaEnxuta)
    If lErro <> SUCESSO Then gError 207582
    NaturezaAte.PromptInclude = False
    NaturezaAte.Text = sNaturezaEnxuta
    NaturezaAte.PromptInclude = True
    Call NaturezaAte_Validate(bSGECancelDummy)
    
    lErro = objRelOpcoes.ObterParametro("NPROJETO", sParam)
    If lErro <> SUCESSO Then gError 194317
    
    objProjeto.lNumIntDoc = StrParaLong(sParam)
    
    If objProjeto.lNumIntDoc <> 0 Then
    
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194318
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194319
        
    End If
    
    glNumIntPRJ = objProjeto.lNumIntDoc
    
    lErro = objRelOpcoes.ObterParametro("NETAPA", sParam)
    If lErro <> SUCESSO Then gError 194320
    
    objEtapa.lNumIntDoc = StrParaLong(sParam)
    
    If objEtapa.lNumIntDoc <> 0 Then
    
        lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194321
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194322
        
    End If
    
    glNumIntPRJEtapa = objEtapa.lNumIntDoc
    
    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 194323
    
    sProjetoAnt = Projeto.Text
    
    sEtapaAnt = objEtapa.sCodigo
    
    Call gobjTelaProjetoInfo.Trata_Etapa(glNumIntPRJ, Etapa)
    
    Call CF("SCombo_Seleciona2", Etapa, sEtapaAnt)

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 194284 To 194286, 194317, 194318, 194320, 194321, 194323
        
        Case 194319
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO", gErr, objProjeto.lNumIntDoc)
        
        Case 194322
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO", gErr, objEtapa.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194287)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 194288

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 194289

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 194288
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 194289
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194290)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 194291
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 194291
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194292)

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
    If ComboOpcoes.Text = "" Then gError 194293

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 194294

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 194295

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 194293
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 194294, 194295
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194296)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    glNumIntPRJ = 0
    glNumIntPRJEtapa = 0
            
    Etapa.Clear
        
    sProjetoAnt = ""
    sEtapaAnt = ""
    
    LabelNaturezaDeDesc.Caption = ""
    LabelNaturezaAteDesc.Caption = ""

    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set objEventoNaturezaDe = New AdmEvento
    Set objEventoNaturezaAte = New AdmEvento
    
    lErro = Inicializa_Mascara_Projeto(Projeto)
    If lErro <> SUCESSO Then gError 194831

    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then gError 194831

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 194831

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194297)

    End Select
   
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRODUTOS
    Set Form_Load_Ocx = Me
    Caption = "Projetos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpFluxoPRJ"
    
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

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(Trim(DataFinal.ClipText)) <> 0 Then

        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 194298

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 194298

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194299)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(Trim(DataInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 194300

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 194300

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194301)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        End If
                
    End If

End Sub

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

Private Sub UpDownDataFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 194302

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 194302

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194303)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_UpClick

    DataFinal.SetFocus

    If Len(Trim(DataFinal.ClipText)) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 194304

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 194304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194305)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicial_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_DownClick

    DataInicial.SetFocus

    If Len(DataInicial.ClipText) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 194306

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 194306

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194307)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicial_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_UpClick

    DataInicial.SetFocus

    If Len(Trim(DataInicial.ClipText)) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 194308

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 194308

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194309)

    End Select

    Exit Sub

End Sub

Sub BotaoProjetos_Click()
    Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Sub

Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Sub

Sub Projeto_GotFocus()
    Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Sub

Sub Projeto_Validate(Cancel As Boolean)
    Call ProjetoTela_Validate(Cancel)
End Sub

Sub Etapa_Validate(Cancel As Boolean)
    Call ProjetoTela_Validate(Cancel)
End Sub

Public Function ProjetoTela_Validate(Cancel As Boolean) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim colItensPRJCR As New Collection
Dim objItemPRJCR As New ClassItensPRJCR
Dim objPRJCR As ClassPRJCR
Dim colPRJCR As New Collection
Dim bPossuiDocOriginal As Boolean
Dim objNF As New ClassNFiscal
Dim objEtapa As New ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_ProjetoTela_Validate

    'Se alterou o projeto
    If sProjetoAnt <> Projeto.Text Or sEtapaAnt <> SCodigo_Extrai(Etapa.Text) Then

        If Len(Trim(Projeto.ClipText)) > 0 Then
                
            lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError 194310
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194311
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194312
            
            If sProjetoAnt <> Projeto.Text Then
                Call gobjTelaProjetoInfo.Trata_Etapa(objProjeto.lNumIntDoc, Etapa)
            End If
            
            If Len(Trim(Etapa.Text)) > 0 Then
            
                objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objEtapa.sCodigo = SCodigo_Extrai(Etapa.Text)
            
                lErro = CF("PrjEtapas_Le", objEtapa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194313
            
            End If
                          
            glNumIntPRJ = objProjeto.lNumIntDoc
            glNumIntPRJEtapa = objEtapa.lNumIntDoc
            
        Else
        
            glNumIntPRJ = 0
            glNumIntPRJEtapa = 0
            
            Etapa.Clear
            
        End If
        
        sProjetoAnt = Projeto.Text
        sEtapaAnt = SCodigo_Extrai(Etapa.Text)
        
    End If
    
    ProjetoTela_Validate = SUCESSO
    
    Exit Function

Erro_ProjetoTela_Validate:

    ProjetoTela_Validate = gErr

    Cancel = True

    Select Case gErr
    
        Case 194310, 194311, 194313
        
        Case 194312
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 194314)

    End Select

    Exit Function

End Function


Private Sub objEventoNaturezaDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaDeEnxuta As String

On Error GoTo Erro_objEventoNaturezaDe_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        NaturezaDe.PromptInclude = False
        NaturezaDe.Text = ""
        NaturezaDe.PromptInclude = True
    
    Else

        sNaturezaDeEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaDeEnxuta)
        If lErro <> SUCESSO Then gError 122833

        NaturezaDe.PromptInclude = False
        NaturezaDe.Text = sNaturezaDeEnxuta
        NaturezaDe.PromptInclude = True
    
    End If

    Call NaturezaDe_Validate(bSGECancelDummy)
    
    'Me.Show

    Exit Sub

Erro_objEventoNaturezaDe_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub NaturezaDe_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaDeFormatada As String
Dim iNaturezaDePreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_NaturezaDe_Validate

    If Len(NaturezaDe.ClipText) > 0 Then

        sNaturezaDeFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da NaturezaDe
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, NaturezaDe.Text, sNaturezaDeFormatada, iNaturezaDePreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaDeFormatada
        
        'Verifica se a NaturezaDe é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_INDEFINIDA)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da NaturezaDe na Tela
        LabelNaturezaDeDesc.Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDeDesc.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_NaturezaDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub


Private Sub LabelNaturezaDe_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = NaturezaDe.ClipText
    
    'colSelecao.Add NaturezaDe_TIPO_PAGAMENTO
    
    'Call Chama_Tela_Modal("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNaturezaDe, "Tipo = ?")
    
    Call Chama_Tela_Modal("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNaturezaDe, "")

End Sub

Private Sub objEventoNaturezaAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaAteEnxuta As String

On Error GoTo Erro_objEventoNaturezaAte_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        NaturezaAte.PromptInclude = False
        NaturezaAte.Text = ""
        NaturezaAte.PromptInclude = True
    
    Else

        sNaturezaAteEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaAteEnxuta)
        If lErro <> SUCESSO Then gError 122833

        NaturezaAte.PromptInclude = False
        NaturezaAte.Text = sNaturezaAteEnxuta
        NaturezaAte.PromptInclude = True
    
    End If

    Call NaturezaAte_Validate(bSGECancelDummy)
    
    'Me.Show

    Exit Sub

Erro_objEventoNaturezaAte_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub NaturezaAte_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaAteFormatada As String
Dim iNaturezaAtePreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_NaturezaAte_Validate

    If Len(NaturezaAte.ClipText) > 0 Then

        sNaturezaAteFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da NaturezaAte
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, NaturezaAte.Text, sNaturezaAteFormatada, iNaturezaAtePreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaAteFormatada
        
        'Verifica se a NaturezaAte é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_INDEFINIDA)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da NaturezaAte na Tela
        LabelNaturezaAteDesc.Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaAteDesc.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_NaturezaAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub


Private Sub LabelNaturezaAte_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = NaturezaAte.ClipText
    
    'colSelecao.Add NaturezaAte_TIPO_PAGAMENTO
    
    'Call Chama_Tela_Modal("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNaturezaAte, "Tipo = ?")
    
    Call Chama_Tela_Modal("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNaturezaAte, "")

End Sub

Private Function Inicializa_Mascara_Natureza() As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    NaturezaDe.Mask = sMascaraNatureza
    NaturezaAte.Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function
