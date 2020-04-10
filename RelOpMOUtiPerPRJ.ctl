VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpMOUtiPerPRJ 
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   6285
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   705
      Left            =   105
      TabIndex        =   26
      Top             =   4035
      Width           =   6045
      Begin VB.OptionButton OptPorData 
         Caption         =   "Por data"
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
         Left            =   780
         TabIndex        =   9
         Top             =   300
         Width           =   1545
      End
      Begin VB.OptionButton OptComp 
         Caption         =   "Comparativo"
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
         Left            =   3495
         TabIndex        =   10
         Top             =   270
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mão de obra"
      Height          =   1332
      Left            =   105
      TabIndex        =   23
      Top             =   2670
      Width           =   6045
      Begin MSMask.MaskEdBox MOFinal 
         Height          =   315
         Left            =   765
         TabIndex        =   8
         Top             =   855
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MOInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   7
         Top             =   330
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelMOAte 
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
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   900
         Width           =   360
      End
      Begin VB.Label LabelMODe 
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   375
         Width           =   315
      End
   End
   Begin VB.ComboBox Etapa 
      Height          =   315
      Left            =   3585
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2250
      Width           =   2550
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   1395
      Left            =   120
      TabIndex        =   18
      Top             =   690
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   855
         Width           =   330
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMOUtiPerPRJ.ctx":0000
      Left            =   840
      List            =   "RelOpMOUtiPerPRJ.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
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
      Picture         =   "RelOpMOUtiPerPRJ.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   825
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4020
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpMOUtiPerPRJ.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpMOUtiPerPRJ.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpMOUtiPerPRJ.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpMOUtiPerPRJ.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Projeto 
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   2265
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
      Left            =   2925
      TabIndex        =   22
      Top             =   2295
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
      TabIndex        =   21
      Top             =   2310
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
      TabIndex        =   17
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpMOUtiPerPRJ"
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

Private WithEvents objEventoMODe As AdmEvento
Attribute objEventoMODe.VB_VarHelpID = -1
Private WithEvents objEventoMOAte As AdmEvento
Attribute objEventoMOAte.VB_VarHelpID = -1

Dim glNumIntPRJ As Long
Dim glNumIntPRJEtapa As Long
                   
Dim sProjetoAnt As String
Dim sEtapaAnt As String

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoMODe = Nothing
    Set objEventoMOAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 194757
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 194758
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 194757
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 194758
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194759)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção
'recebe os MOs inicial e final no formato do BD

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

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 194760)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(iTipo As Integer) As Long

Dim lErro As Long
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    If StrParaDate(DataInicial.Text) <> DATA_NULA And StrParaDate(DataFinal.Text) <> DATA_NULA Then
        If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 194761
    End If
    
    'se ambos os MOs estão preenchidos, o MO inicial não pode ser maior que o final
    If Codigo_Extrai(MOInicial.Text) And Codigo_Extrai(MOFinal.Text) Then

        If Codigo_Extrai(MOInicial.Text) > Codigo_Extrai(MOFinal.Text) Then gError 194762

    End If
    
    If OptPorData.Value Then
        iTipo = MARCADO
    Else
        iTipo = DESMARCADO
    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 194761
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus

        Case 194762
            Call Rotina_Erro(vbOKOnly, "ERRO_MO_INICIAL_MAIOR", gErr)
            MOInicial.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194763)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lNumIntRel As Long
Dim lNumIntDocProjeto As Long
Dim lNumIntDocEtapa As Long
Dim iTipo As Integer

On Error GoTo Erro_PreencherRelOp
                      
    lErro = Formata_E_Critica_Parametros(iTipo)
    If lErro <> SUCESSO Then gError 194764

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 194765
    
    lErro = objRelOpcoes.IncluirParametro("DDATAINI", Format(StrParaDate(DataInicial.Text), "dd/mm/yyyy"))
    If lErro <> AD_BOOL_TRUE Then gError 194766
    
    lErro = objRelOpcoes.IncluirParametro("DDATAFIM", Format(StrParaDate(DataFinal.Text), "dd/mm/yyyy"))
    If lErro <> AD_BOOL_TRUE Then gError 194767

    lErro = objRelOpcoes.IncluirParametro("TPROJETO", Projeto.Text)
    If lErro <> AD_BOOL_TRUE Then gError 194768

    lErro = objRelOpcoes.IncluirParametro("TETAPA", Etapa.Text)
    If lErro <> AD_BOOL_TRUE Then gError 194769

    lErro = objRelOpcoes.IncluirParametro("NPROJETO", CStr(glNumIntPRJ))
    If lErro <> AD_BOOL_TRUE Then gError 194770

    lErro = objRelOpcoes.IncluirParametro("NETAPA", CStr(glNumIntPRJEtapa))
    If lErro <> AD_BOOL_TRUE Then gError 194771
    
    lErro = objRelOpcoes.IncluirParametro("NMOINIC", Str(Codigo_Extrai(MOInicial.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 194772

    lErro = objRelOpcoes.IncluirParametro("NMOFIM", Str(Codigo_Extrai(MOFinal.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 194773
    
    lErro = objRelOpcoes.IncluirParametro("TMOINIC", MOInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 194774

    lErro = objRelOpcoes.IncluirParametro("TMOFIM", MOFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 194775
    
    lErro = objRelOpcoes.IncluirParametro("NTIPO", CStr(iTipo))
    If lErro <> AD_BOOL_TRUE Then gError 194776
    
    If bExecutando Then
    
        lErro = CF("RelMOUtiPerPRJ_Prepara", lNumIntRel, glNumIntPRJ, glNumIntPRJEtapa, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), Codigo_Extrai(MOInicial.Text), Codigo_Extrai(MOFinal.Text), iTipo)
        If lErro <> SUCESSO Then gError 194777
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 194778
    
    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 194764 To 194778
            'erro tratado nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194779)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim objProjeto As New ClassProjetos
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 194780

    'pega a Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINI", sParam)
    If lErro <> SUCESSO Then gError 194781
    Call DateParaMasked(DataInicial, StrParaDate(sParam))
    
    'pega a Data Final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 194782
    Call DateParaMasked(DataFinal, StrParaDate(sParam))
    
    lErro = objRelOpcoes.ObterParametro("NTIPO", sParam)
    If lErro <> SUCESSO Then gError 194783
    
    If StrParaInt(sParam) = MARCADO Then
        OptPorData.Value = True
    Else
        OptComp.Value = True
    End If
    
    lErro = objRelOpcoes.ObterParametro("NPROJETO", sParam)
    If lErro <> SUCESSO Then gError 194784
    
    objProjeto.lNumIntDoc = StrParaLong(sParam)
    
    If objProjeto.lNumIntDoc <> 0 Then
    
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194785
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194786
        
    End If
    
    glNumIntPRJ = objProjeto.lNumIntDoc
    
    lErro = objRelOpcoes.ObterParametro("NETAPA", sParam)
    If lErro <> SUCESSO Then gError 194787
    
    objEtapa.lNumIntDoc = StrParaLong(sParam)
    
    If objEtapa.lNumIntDoc <> 0 Then
    
        lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194788
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194789
        
    End If
    
    glNumIntPRJEtapa = objEtapa.lNumIntDoc
    
    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 194790
    
    sProjetoAnt = Projeto.Text
    
    sEtapaAnt = objEtapa.sCodigo
    
    Call gobjTelaProjetoInfo.Trata_Etapa(glNumIntPRJ, Etapa)
    
    Call CF("SCombo_Seleciona2", Etapa, sEtapaAnt)
    
    'pega MO Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NMOINIC", sParam)
    If lErro <> SUCESSO Then gError 194791

    If StrParaInt(sParam) <> 0 Then
        MOInicial.Text = StrParaInt(sParam)
        Call MOInicial_Validate(bSGECancelDummy)
    End If

    'pega parâmetro MO Final e exibe
    lErro = objRelOpcoes.ObterParametro("NMOFIM", sParam)
    If lErro <> SUCESSO Then gError 194792

    If StrParaInt(sParam) <> 0 Then
        MOFinal.Text = StrParaInt(sParam)
        Call MOFinal_Validate(bSGECancelDummy)
    End If

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 194780 To 194785, 194787, 194788, 194790 To 194792
        
        Case 194786
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO", gErr, objProjeto.lNumIntDoc)
        
        Case 194789
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO", gErr, objEtapa.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194793)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 194794

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 194795

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 194794
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 194795
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194796)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 194797
    
    If OptComp.Value Then
        gobjRelatorio.sNomeTsk = "CpMOPRJ"
    End If
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 194797
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194798)

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
    If ComboOpcoes.Text = "" Then gError 194799

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 194800

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 194801

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 194799
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 194800, 194801
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194802)

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
    
    OptPorData.Value = True

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

    Set objEventoMODe = New AdmEvento
    Set objEventoMOAte = New AdmEvento

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    OptPorData.Value = True
    
    lErro = Inicializa_Mascara_Projeto(Projeto)
    If lErro <> SUCESSO Then gError 194833

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 194833

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194803)

    End Select
   
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRODUTOS
    Set Form_Load_Ocx = Me
    Caption = "Mãos de obra utilizadas no período"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpMOUtiPerPRJ"
    
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
        If lErro <> SUCESSO Then gError 194804

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 194804

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194805)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(Trim(DataInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 194806

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 194806

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194807)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        ElseIf Me.ActiveControl Is MOInicial Then
            Call LabelMODe_Click
        ElseIf Me.ActiveControl Is MOFinal Then
            Call LabelMOAte_Click
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
        If lErro <> SUCESSO Then gError 194808

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 194808

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194809)

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
        If lErro <> SUCESSO Then gError 194810

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 194810

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194811)

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
        If lErro <> SUCESSO Then gError 194812

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 194812

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194813)

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
        If lErro <> SUCESSO Then gError 194814

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 194814

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194815)

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
            If lErro <> SUCESSO Then gError 194816
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194817
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194818
            
            If sProjetoAnt <> Projeto.Text Then
                Call gobjTelaProjetoInfo.Trata_Etapa(objProjeto.lNumIntDoc, Etapa)
            End If
            
            If Len(Trim(Etapa.Text)) > 0 Then
            
                objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objEtapa.sCodigo = SCodigo_Extrai(Etapa.Text)
            
                lErro = CF("PrjEtapas_Le", objEtapa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194819
            
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
    
        Case 194816, 194817, 194819
        
        Case 194818
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 194820)

    End Select

    Exit Function

End Function

Private Sub LabelMOAte_Click()

Dim lErro As Long
Dim objMO As New ClassTiposDeMaodeObra
Dim colSelecao As New Collection

On Error GoTo Erro_LabelMOAte_Click

    objMO.iCodigo = Codigo_Extrai(MOFinal.Text)

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objMO, objEventoMOAte)

    Exit Sub

Erro_LabelMOAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194821)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMODe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMO As ClassTiposDeMaodeObra


On Error GoTo Erro_objEventoMODe_evSelecao

    Set objMO = obj1

    MOInicial.Text = objMO.iCodigo
    Call MOInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoMODe_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194822)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMOAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMO As ClassTiposDeMaodeObra


On Error GoTo Erro_objEventoMOAte_evSelecao

    Set objMO = obj1

    MOFinal.Text = objMO.iCodigo
    Call MOFinal_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoMOAte_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194823)

    End Select

    Exit Sub

End Sub

Private Sub LabelMODe_Click()
    
Dim lErro As Long
Dim objMO As New ClassTiposDeMaodeObra
Dim colSelecao As New Collection

On Error GoTo Erro_LabelMODe_Click

    objMO.iCodigo = Codigo_Extrai(MOInicial.Text)

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objMO, objEventoMODe)

    Exit Sub

Erro_LabelMODe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194824)

    End Select

    Exit Sub
    
End Sub

Private Sub MOFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMO As New ClassTiposDeMaodeObra

On Error GoTo Erro_MOFinal_Validate
        
    If Codigo_Extrai(MOFinal.Text) <> 0 Then
        
        objMO.iCodigo = Codigo_Extrai(MOFinal.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objMO)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 194825
            
        If lErro <> SUCESSO Then gError 194826
        
        MOFinal.Text = CStr(objMO.iCodigo) & SEPARADOR & objMO.sDescricao
        
    End If
            
    Exit Sub

Erro_MOFinal_Validate:

    Cancel = True

    Select Case gErr
    
        Case 194825
    
        Case 194826
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objMO.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194827)

    End Select

    Exit Sub

End Sub

Private Sub MOInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMO As New ClassTiposDeMaodeObra

On Error GoTo Erro_MOFinal_Validate

    If Codigo_Extrai(MOInicial.Text) <> 0 Then
        
        objMO.iCodigo = Codigo_Extrai(MOInicial.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objMO)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 194828
            
        If lErro <> SUCESSO Then gError 194829
        
        MOInicial.Text = CStr(objMO.iCodigo) & SEPARADOR & objMO.sDescricao
        
    End If
        
    Exit Sub

Erro_MOFinal_Validate:

    Cancel = True

    Select Case gErr
    
        Case 194828
    
        Case 194829
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objMO.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194830)

    End Select

    Exit Sub

End Sub

Private Sub LabelMODe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMODe, Source, X, Y)
End Sub

Private Sub LabelMODe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMODe, Button, Shift, X, Y)
End Sub

Private Sub LabelMOAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMOAte, Source, X, Y)
End Sub

Private Sub LabelMOAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMOAte, Button, Shift, X, Y)
End Sub
