VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpUsoMaquinaOcx 
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LockControls    =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   8130
   Begin VB.Frame FrameMaquinas 
      Caption         =   "Máquinas"
      Height          =   825
      Left            =   90
      TabIndex        =   24
      Top             =   3210
      Width           =   7935
      Begin VB.TextBox MaquinaInicial 
         Height          =   300
         Left            =   450
         MaxLength       =   26
         TabIndex        =   7
         Top             =   300
         Width           =   2895
      End
      Begin VB.TextBox MaquinaFinal 
         Height          =   300
         Left            =   4875
         MaxLength       =   26
         TabIndex        =   8
         Top             =   300
         Width           =   2895
      End
      Begin VB.Label LabelMaquinaInicial 
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   360
         Width           =   375
      End
      Begin VB.Label LabelMaquinaFinal 
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
         Left            =   4500
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame FrameCT 
      Caption         =   "Centros de Trabalho"
      Height          =   1395
      Left            =   90
      TabIndex        =   19
      Top             =   1710
      Width           =   7935
      Begin MSMask.MaskEdBox CTInicial 
         Height          =   315
         Left            =   525
         TabIndex        =   5
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTFinal 
         Height          =   315
         Left            =   525
         TabIndex        =   6
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescCTFinal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2085
         TabIndex        =   23
         Top             =   840
         Width           =   5640
      End
      Begin VB.Label DescCTInicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2085
         TabIndex        =   22
         Top             =   360
         Width           =   5640
      End
      Begin VB.Label LabelCTDe 
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
         Height          =   255
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   390
         Width           =   360
      End
      Begin VB.Label LabelCTAte 
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
         Height          =   255
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   885
         Width           =   435
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   810
      Left            =   90
      TabIndex        =   16
      Top             =   780
      Width           =   7935
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   525
         TabIndex        =   1
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataInicial 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   4875
         TabIndex        =   3
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataFinal 
         Height          =   300
         Left            =   6045
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   330
         Width           =   315
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4500
         TabIndex        =   17
         Top             =   330
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpUsoMaquina.ctx":0000
      Left            =   840
      List            =   "RelOpUsoMaquina.ctx":0002
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
      Left            =   4080
      Picture         =   "RelOpUsoMaquina.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5865
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpUsoMaquina.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpUsoMaquina.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpUsoMaquina.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpUsoMaquina.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
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
      Left            =   135
      TabIndex        =   15
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpUsoMaquinaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giMaquinaInicial As Integer

Private WithEvents objEventoCTInic As AdmEvento
Attribute objEventoCTInic.VB_VarHelpID = -1
Private WithEvents objEventoCTFim As AdmEvento
Attribute objEventoCTFim.VB_VarHelpID = -1
Private WithEvents objEventoMaquinaDe As AdmEvento
Attribute objEventoMaquinaDe.VB_VarHelpID = -1
Private WithEvents objEventoMaquinaAte As AdmEvento
Attribute objEventoMaquinaAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giFocoInicial As Integer

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoCTInic = Nothing
    Set objEventoCTFim = Nothing
    Set objEventoMaquinaDe = Nothing
    Set objEventoMaquinaAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 141127
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 141128
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 141127
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 141128
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173579)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, lCT_I As Long, lCT_F As Long, lMaq_I As Long, lMaq_F As Long) As Long
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

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173580)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(lCTIni As Long, lCTFim As Long, iMaqIni As Integer, iMaqFim As Integer) As Long

Dim objCTInicial As ClassCentrodeTrabalho
Dim objCTFinal As ClassCentrodeTrabalho
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 141134

    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 141135
    
    If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 141132
    
    Set objCTInicial = New ClassCentrodeTrabalho
    
    objCTInicial.sNomeReduzido = Trim(CTInicial.Text)
    
    'Lê CT Inicial pelo NomeReduzido
    lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCTInicial)
    If lErro <> SUCESSO And lErro <> 134941 Then gError 141129
            
    lCTIni = objCTInicial.lCodigo
            
    Set objCTFinal = New ClassCentrodeTrabalho
    
    objCTFinal.sNomeReduzido = Trim(CTFinal.Text)
    
    'Lê CT Final pelo NomeReduzido
    lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCTFinal)
    If lErro <> SUCESSO And lErro <> 134941 Then gError 141130
            
    lCTFim = objCTFinal.lCodigo
    
    'Valida Centros de Trabalho
    If Len(Trim(CTInicial.Text)) <> 0 And Len(Trim(CTFinal.Text)) <> 0 Then
    
        'codigo do CT inicial não pode ser maior que o final
        If objCTInicial.lCodigo > objCTFinal.lCodigo Then gError 141131
        
    End If
        
    iMaqIni = Codigo_Extrai(MaquinaInicial.Text)
    iMaqFim = Codigo_Extrai(MaquinaFinal.Text)

    'maquina inicial não pode ser maior que a maquina final
    If iMaqIni <> 0 And iMaqFim <> 0 Then
    
        If iMaqIni > iMaqFim Then gError 141133
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 141129, 141130
        
        Case 141131
            Call Rotina_Erro(vbOKOnly, "ERRO_CT_INICIAL_MAIOR", gErr)
            CTInicial.SetFocus
    
        Case 141132
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
        
        Case 141133
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_INICIAL_MAIOR", gErr)
            MaquinaInicial.SetFocus
            
        Case 141134
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus

        Case 141135
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173581)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lCTIni As Long
Dim lCTFim As Long
Dim iMaqIni As Integer
Dim iMaqFim As Integer
Dim lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp
           
    lErro = Formata_E_Critica_Parametros(lCTIni, lCTFim, iMaqIni, iMaqFim)
    If lErro <> SUCESSO Then gError 141136

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 141137

    lErro = objRelOpcoes.IncluirParametro("NCTINIC", CStr(lCTIni))
    If lErro <> AD_BOOL_TRUE Then gError 141138

    lErro = objRelOpcoes.IncluirParametro("NCTFIM", CStr(lCTFim))
    If lErro <> AD_BOOL_TRUE Then gError 141139

    lErro = objRelOpcoes.IncluirParametro("TCTINIC", CTInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141140

    lErro = objRelOpcoes.IncluirParametro("TCTFIM", CTFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141141
    
    lErro = objRelOpcoes.IncluirParametro("NMAQINIC", CStr(iMaqIni))
    If lErro <> AD_BOOL_TRUE Then gError 141142

    lErro = objRelOpcoes.IncluirParametro("NMAQFIM", CStr(iMaqFim))
    If lErro <> AD_BOOL_TRUE Then gError 141143

    lErro = objRelOpcoes.IncluirParametro("TMAQINIC", MaquinaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141144

    lErro = objRelOpcoes.IncluirParametro("TMAQFIM", MaquinaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141145
    
    If Len(Trim(DataInicial.ClipText)) <> 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINI", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 141146
    
    If Len(Trim(DataFinal.ClipText)) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 141147

'    lErro = Monta_Expressao_Selecao(objRelOpcoes, lCTIni, lCTFim)
'    If lErro <> SUCESSO Then gError 137863

    If bExecutando Then
    
        lErro = CF("RelUsoMaquina_Prepara", lNumIntRel, giFilialEmpresa, iMaqIni, iMaqFim, lCTIni, lCTFim, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text))
        If lErro <> SUCESSO Then gError 141148
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 141149
    
    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 141136 To 141149
            'erro tratado nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173582)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 141150

    'pega CT Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCTINIC", sParam)
    If lErro Then gError 141151

    If Len(Trim(sParam)) > 0 Then
 
        CTInicial.Text = sParam
        Call CTInicial_Validate(bSGECancelDummy)
        
    End If

    'pega CT Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCTFIM", sParam)
    If lErro Then gError 141152

    If Len(Trim(sParam)) > 0 Then
 
        CTFinal.Text = sParam
        Call CTFinal_Validate(bSGECancelDummy)
        
    End If

    'pega a Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINI", sParam)
    If lErro <> SUCESSO Then gError 141153
    Call DateParaMasked(DataInicial, StrParaDate(sParam))
    
    'pega a Data Final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 141154
    Call DateParaMasked(DataFinal, StrParaDate(sParam))

    'pega a máquina inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NMAQINIC", sParam)
    If lErro <> SUCESSO Then gError 141155
    
    MaquinaInicial.Text = sParam
    Call MaquinaInicial_Validate(bSGECancelDummy)

    'pega a máquina final e exibe
    lErro = objRelOpcoes.ObterParametro("NMAQFIM", sParam)
    If lErro <> SUCESSO Then gError 141156
    
    MaquinaFinal.Text = sParam
    Call MaquinaFinal_Validate(bSGECancelDummy)

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 141150 To 141156
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173583)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 141157

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 141158

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 141157
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 141158
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173584)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 141159
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 141159
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173585)

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
    If ComboOpcoes.Text = "" Then gError 141160

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 141161

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 141162

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 141160
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 141161, 141162
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173586)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    DescCTInicial.Caption = ""
    DescCTFinal.Caption = ""

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

On Error GoTo Erro_RelOpProdutos_Form_Load
    
    giFocoInicial = 1
    
    Set objEventoCTInic = New AdmEvento
    Set objEventoCTFim = New AdmEvento
    Set objEventoMaquinaDe = New AdmEvento
    Set objEventoMaquinaAte = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_RelOpProdutos_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173587)

    End Select
   
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRODUTOS
    Set Form_Load_Ocx = Me
    Caption = "Relação de Ordens de Trabalho"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpOrdensDeTrabalho"
    
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
        If lErro <> SUCESSO Then gError 141163

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 141163

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173588)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(Trim(DataInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 141164

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 141164

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173589)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is CTInicial Then
            Call LabelCTDe_Click
        ElseIf Me.ActiveControl Is CTFinal Then
            Call LabelCTAte_Click
        ElseIf Me.ActiveControl Is MaquinaInicial Then
            Call LabelMaquinaInicial_Click
        ElseIf Me.ActiveControl Is MaquinaFinal Then
            Call LabelMaquinaFinal_Click
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

Private Sub CTFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_CTFinal_Validate

    DescCTFinal.Caption = ""

    'Verifica se CTFinal não está preenchido
    If Len(Trim(CTFinal.Text)) <> 0 Then
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CTFinal, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 141165
                
        DescCTFinal.Caption = objCentrodeTrabalho.sDescricao
           
    End If
    
    Exit Sub

Erro_CTFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 141165
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173590)

    End Select

    Exit Sub

End Sub

Private Sub CTInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_CTInicial_Validate

    DescCTInicial.Caption = ""

    'Verifica se CTInicial não está preenchido
    If Len(Trim(CTInicial.Text)) <> 0 Then

        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CTInicial, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 141166
                
        DescCTInicial.Caption = objCentrodeTrabalho.sDescricao
       
    End If
    
    Exit Sub

Erro_CTInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 141166
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173591)

    End Select

    Exit Sub

End Sub

Private Sub LabelCTAte_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCTAte

    'Verifica se o CTFinal foi preenchido
    If Len(Trim(CTFinal.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = Trim(CTFinal.Text)
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCTFim)

    Exit Sub

Erro_LabelCTAte:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173592)

    End Select

    Exit Sub

End Sub

Private Sub LabelCTDe_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCTDe

    'Verifica se o CTInicial foi preenchido
    If Len(Trim(CTInicial.Text)) <> 0 Then
    
        objCentrodeTrabalho.sNomeReduzido = Trim(CTInicial.Text)
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCTInic)

    Exit Sub

Erro_LabelCTDe:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173593)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCTFim_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCTFim_evSelecao

    Set objCentrodeTrabalho = obj1

    CTFinal.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CTFinal_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCTFim_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173594)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCTInic_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCTInic_evSelecao

    Set objCentrodeTrabalho = obj1

    CTInicial.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CTInicial_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCTInic_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173595)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 141167

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 141167

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173596)

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
        If lErro <> SUCESSO Then gError 141168

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 141168

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173597)

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
        If lErro <> SUCESSO Then gError 141169

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 141169

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173598)

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
        If lErro <> SUCESSO Then gError 141170

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 141170

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173599)

    End Select

    Exit Sub

End Sub

Private Sub LabelMaquinaFinal_Click()

Dim lErro As Long
Dim objMaquina As New ClassMaquinas
Dim colSelecao As Collection

On Error GoTo Erro_LabelMaquinaFinal_Click

    giMaquinaInicial = 0

    If Len(Trim(MaquinaFinal.Text)) <> 0 Then

        objMaquina.sNomeReduzido = MaquinaFinal.Text

    End If

    Call Chama_Tela("MaquinasLista", colSelecao, objMaquina, objEventoMaquinaAte)

    Exit Sub

Erro_LabelMaquinaFinal_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173600)

    End Select

End Sub

Private Sub LabelMaquinaInicial_Click()

Dim lErro As Long
Dim objMaquina As New ClassMaquinas
Dim colSelecao As Collection

On Error GoTo Erro_LabelMaquinaInicial_Click

    giMaquinaInicial = 1

    If Len(Trim(MaquinaInicial.Text)) <> 0 Then

        objMaquina.sNomeReduzido = MaquinaInicial.Text

    End If

    Call Chama_Tela("MaquinasLista", colSelecao, objMaquina, objEventoMaquinaDe)

    Exit Sub

Erro_LabelMaquinaInicial_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173601)

    End Select

End Sub

Private Sub MaquinaFinal_GotFocus()

    giMaquinaInicial = 0

End Sub

Private Sub MaquinaInicial_GotFocus()

    giMaquinaInicial = 1

End Sub

Private Sub MaquinaFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_MaquinaFinal_Validate

    If Len(Trim(MaquinaFinal.Text)) > 0 Then
        
        lErro = CF("TP_Maquina_Le2", MaquinaFinal, objMaquina)
        If lErro <> SUCESSO And lErro <> 106451 And lErro <> 106453 Then gError 141171
        
        'Se nao encontrou => Erro
        If lErro = 106451 Or lErro = 106453 Then gError 141172
        
    End If

    Exit Sub

Erro_MaquinaFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 141171
        
        Case 141172
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, MaquinaFinal.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173602)

    End Select

End Sub

Private Sub MaquinaInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_MaquinaInicial_Validate

    If Len(Trim(MaquinaInicial.Text)) > 0 Then
        
        lErro = CF("TP_Maquina_Le2", MaquinaInicial, objMaquina)
        If lErro <> SUCESSO And lErro <> 106451 And lErro <> 106453 Then gError 141173
        
        'Se nao encontrou => Erro
        If lErro = 106451 Or lErro = 106453 Then gError 141174
        
    End If

    Exit Sub

Erro_MaquinaInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 141173
        
        Case 141174
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, MaquinaInicial.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173603)

    End Select

End Sub

Private Sub objEventoMaquinaAte_evSelecao(obj1 As Object)
'Evento do Browser

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_objEventoMaquinaAte_evSelecao

    Set objMaquina = obj1
    
    objMaquina.iFilialEmpresa = giFilialEmpresa
    
    'Tenta Ler a maquina
    lErro = CF("Maquinas_Le", objMaquina)
    If lErro <> SUCESSO And lErro <> 103090 Then gError 141175
    
    'Se nao Encontrou => Erro
    If lErro = 103090 Then gError 141176
    
    'Coloca na Tela o Codigo "-" NomeReduzido
    MaquinaFinal.Text = objMaquina.iCodigo & SEPARADOR & objMaquina.sNomeReduzido
    
    Me.Show
    
    Exit Sub

Erro_objEventoMaquinaAte_evSelecao:

    Select Case gErr
    
        Case 141175
        
        Case 141176
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173604)
            
    End Select

End Sub

Private Sub objEventoMaquinaDe_evSelecao(obj1 As Object)
'Evento do Browser

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_objEventoMaquinaDe_evSelecao

    Set objMaquina = obj1
    
    objMaquina.iFilialEmpresa = giFilialEmpresa
    
    'Tenta Ler a maquina
    lErro = CF("Maquinas_Le", objMaquina)
    If lErro <> SUCESSO And lErro <> 103090 Then gError 141177
    
    'Se nao Encontrou => Erro
    If lErro = 103090 Then gError 141178
    
    'Coloca na Tela o Codigo "-" NomeReduzido
    MaquinaInicial.Text = objMaquina.iCodigo & SEPARADOR & objMaquina.sNomeReduzido
    
    Me.Show
    
    Exit Sub

Erro_objEventoMaquinaDe_evSelecao:

    Select Case gErr
    
        Case 141177
        
        Case 141178
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173605)
            
    End Select

End Sub
