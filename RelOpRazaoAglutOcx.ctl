VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpRazaoAglutOcx 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8385
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   8385
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
      Height          =   615
      Left            =   4245
      Picture         =   "RelOpRazaoAglutOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRazaoAglutOcx.ctx":0102
      Left            =   1395
      List            =   "RelOpRazaoAglutOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   360
      Width           =   2610
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   240
      TabIndex        =   5
      Top             =   1530
      Width           =   7995
      Begin VB.CheckBox CheckPulaPag 
         Caption         =   "Pula p�gina a cada nova conta"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1410
         Width           =   3015
      End
      Begin MSMask.MaskEdBox ContaFinal 
         Height          =   315
         Left            =   690
         TabIndex        =   7
         Top             =   930
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaInicial 
         Height          =   315
         Left            =   690
         TabIndex        =   8
         Top             =   330
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label DescCtaFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2685
         TabIndex        =   12
         Top             =   930
         Width           =   4995
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
         Height          =   255
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.Label DescCtaInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2685
         TabIndex        =   10
         Top             =   330
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
         Height          =   255
         Left            =   75
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6120
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRazaoAglutOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRazaoAglutOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRazaoAglutOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRazaoAglutOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   2745
      Left            =   5940
      TabIndex        =   15
      Top             =   1155
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
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   2550
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1020
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   315
      Left            =   1395
      TabIndex        =   17
      Top             =   1050
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   330
      Left            =   5145
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1035
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   315
      Left            =   3990
      TabIndex        =   19
      Top             =   1050
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
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
      Left            =   5925
      TabIndex        =   23
      Top             =   930
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial:"
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
      Left            =   270
      TabIndex        =   22
      Top             =   1155
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
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
      Left            =   2970
      TabIndex        =   21
      Top             =   1110
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Op��o:"
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
      Left            =   690
      TabIndex        =   20
      Top             =   420
      Width           =   630
   End
End
Attribute VB_Name = "RelOpRazaoAglutOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giFocoInicial As Boolean
Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Private WithEvents objEventoContaDe As AdmEvento
Attribute objEventoContaDe.VB_VarHelpID = -1
Private WithEvents objEventoContaAte As AdmEvento
Attribute objEventoContaAte.VB_VarHelpID = -1

Function Obtem_Periodo_Exercicio(iPer_I As Integer, iPer_F As Integer, iExercicio As Integer, sDtIni_I As String, sDtFim_F As String) As Long
'a partir das datas ( inicial e final ) encontra o per�odo e o exerc�cio
'as datas devem estar no mesmo exerc�cio
'pega tamb�m a data inicial do per�odo inicial e a data final do per�odo final

Dim objPer_I As New ClassPeriodo, objPer_F As New ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_Obtem_Periodo_Exercicio

    'pega o per�odo da Data Inicial
    lErro = CF("Periodo_Le", DataInicial.Text, objPer_I)
    If lErro <> SUCESSO Then Error 65303

    'pega o per�odo da Data Final
    lErro = CF("Periodo_Le", DataFinal.Text, objPer_F)
    If lErro <> SUCESSO Then Error 65304

    'Data Inicial e Final devem estar num mesmo exerc�cio
    If objPer_I.iExercicio <> objPer_F.iExercicio Then Error 65305

    iPer_I = objPer_I.iPeriodo
    iPer_F = objPer_F.iPeriodo
    iExercicio = objPer_I.iExercicio
    sDtIni_I = objPer_I.dtDataInicio
    sDtFim_F = objPer_I.dtDataFim

    Obtem_Periodo_Exercicio = SUCESSO

    Exit Function

Erro_Obtem_Periodo_Exercicio:

    Obtem_Periodo_Exercicio = Err

    Select Case Err

        Case 65303
            DataInicial.SetFocus

        Case 65304
            DataFinal.SetFocus

        Case 65305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_COM_EXERCICIOS_DIFERENTES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171978)

    End Select

    Exit Function

End Function

Function Formata_E_Critica_Contas(sCta_I As String, sCta_F As String) As Long
'retorna em sCta_I e sCta_F as contas (inicial e final) formatadas
'Verifica se a conta inicial � maior que a conta final

Dim iCtaPreenchida_I As Integer, iCtaPreenchida_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Contas

    'formata a Conta Inicial
    lErro = CF("Conta_Formata", ContaInicial.Text, sCta_I, iCtaPreenchida_I)
    If lErro <> SUCESSO Then Error 65306
    If iCtaPreenchida_I <> CONTA_PREENCHIDA Then sCta_I = ""

    'formata a Conta Final
    lErro = CF("Conta_Formata", ContaFinal.Text, sCta_F, iCtaPreenchida_F)
    If lErro <> SUCESSO Then Error 65307
    If iCtaPreenchida_F <> CONTA_PREENCHIDA Then sCta_F = ""

    'se ambas as contas est�o preenchidas, a conta inicial n�o pode ser maior que a final
    If iCtaPreenchida_I = CONTA_PREENCHIDA And iCtaPreenchida_F = CONTA_PREENCHIDA Then
        If sCta_I > sCta_F Then Error 65308
    End If

    Formata_E_Critica_Contas = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Contas:

    Formata_E_Critica_Contas = Err

    Select Case Err

        Case 65306
            ContaInicial.SetFocus

        Case 65307
            ContaFinal.SetFocus

        Case 65308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171979)

    End Select

    Exit Function

End Function

Private Function Traz_Conta_Tela(sConta As String) As Long
'verifica e preenche a conta inicial e final com sua descri�ao de acordo com o �ltimo foco
'sConta deve estar no formato do BD

Dim lErro As Long

On Error GoTo Erro_Traz_Conta_Tela

    If giFocoInicial Then

        lErro = CF("Traz_Conta_MaskEd", sConta, ContaInicial, DescCtaInic)
        If lErro <> SUCESSO Then Error 65309

    Else

        lErro = CF("Traz_Conta_MaskEd", sConta, ContaFinal, DescCtaFim)
        If lErro <> SUCESSO Then Error 65310

    End If

    Traz_Conta_Tela = SUCESSO

    Exit Function

Erro_Traz_Conta_Tela:

    Traz_Conta_Tela = Err

    Select Case Err

        Case 65309
        
        Case 65310

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171980)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpRazao() As Long
'faz a cr�tica da data inicial e da data final

Dim lErro As Long

On Error GoTo Erro_Critica_Datas_RelOpRazao

    'data inicial n�o pode ser vazia
    If Len(DataInicial.ClipText) = 0 Then Error 65311

    'data final n�o pode ser vazia
    If Len(DataFinal.ClipText) = 0 Then Error 65312

    'data inicial n�o pode ser maior que a data final
    If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 65313

    Critica_Datas_RelOpRazao = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpRazao:

    Critica_Datas_RelOpRazao = Err

    Select Case Err

        Case 65311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)
            DataInicial.SetFocus

        Case 65312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)
            DataFinal.SetFocus

        Case 65313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171981)

    End Select
    
    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros do arqquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 65314

    'pega Conta Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAINIC", sParam)
    If lErro <> SUCESSO Then Error 65315

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then Error 65316

    'pega Conta Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAFIM", sParam)
    If lErro <> SUCESSO Then Error 65317

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaFinal, DescCtaFim)
    If lErro <> SUCESSO Then Error 65318

    'pega 'Pula p�gina a cada novo conta' e exibe
    lErro = objRelOpcoes.ObterParametro("TPULAPAGQBR0", sParam)
    If lErro <> SUCESSO Then Error 65319

    If sParam = "S" Then CheckPulaPag.Value = 1
    If sParam = "N" Then CheckPulaPag.Value = 0
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 65320

    DataInicial.PromptInclude = False
    DataInicial.Text = sParam
    DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 65321

    DataFinal.PromptInclude = False
    DataFinal.Text = sParam
    DataFinal.PromptInclude = True

    PreencherParametrosNaTela = SUCESSO
    
    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err
    
    Select Case Err

        Case 65314 To 65321

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171982)

    End Select
    
    Exit Function
    
End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bGeraArqTemp As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usu�rio

Dim lErro As Long, lNumIntRel As Long
Dim iPer_I As Integer, iPer_F As Integer
Dim iExercicio As Integer
Dim sCheck As String
Dim sDtIni_I As String, sDtFim_F As String
Dim sCta_I As String, sCta_F As String

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_RelOpRazao
    If lErro <> SUCESSO Then Error 65322

    lErro = Formata_E_Critica_Contas(sCta_I, sCta_F)
    If lErro <> SUCESSO Then Error 65323

    lErro = Obtem_Periodo_Exercicio(iPer_I, iPer_F, iExercicio, sDtIni_I, sDtFim_F)
    If lErro <> SUCESSO Then Error 65324

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 65325

    lErro = objRelOpcoes.IncluirParametro("TCTAINIC", sCta_I)
    If lErro <> AD_BOOL_TRUE Then Error 65326

    lErro = objRelOpcoes.IncluirParametro("TCTAFIM", sCta_F)
    If lErro <> AD_BOOL_TRUE Then Error 65327

    'Pula P�gina a Cada Nova Conta
    If CheckPulaPag.Value Then
        sCheck = "S"
    Else
        sCheck = "N"
    End If

    lErro = objRelOpcoes.IncluirParametro("TPULAPAGQBR0", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 65328

    lErro = objRelOpcoes.IncluirParametro("NPERINIC", CStr(iPer_I))
    If lErro <> AD_BOOL_TRUE Then Error 65329

    lErro = objRelOpcoes.IncluirParametro("NPERFIM", CStr(iPer_F))
    If lErro <> AD_BOOL_TRUE Then Error 65330

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(iExercicio))
    If lErro <> AD_BOOL_TRUE Then Error 65331

    lErro = objRelOpcoes.IncluirParametro("DINICPERINI", sDtIni_I)
    If lErro <> AD_BOOL_TRUE Then Error 65332

    lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65333

    lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65334

    lErro = Maior_Menor_Conta(objRelOpcoes, sDtIni_I)
    If lErro <> SUCESSO Then Error 65335

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCta_I, sCta_F, sDtIni_I, sDtFim_F)
    If lErro <> SUCESSO Then Error 65336

    Call Acha_Nome_TSK(sDtIni_I)
    
    If bGeraArqTemp Then
    
        lErro = CF("RelCtaSaldo_Prepara", giFilialEmpresa, lNumIntRel, sCta_I, sCta_F, CDate(DataInicial.Text), CDate(DataFinal.Text))
        If lErro <> SUCESSO Then Error 54966
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then Error 54964

    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err
    
    Select Case Err

        Case 65322 To 65336

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171983)
            
    End Select
    
    Exit Function
    
End Function

Function Maior_Menor_Conta(objRelOpcoes As AdmRelOpcoes, sDtIni_I As String) As Long
'acha a menor e a maior conta do BD se a data de in�cio n�o coincide com o in�cio do per�odo inicial
'preenche TCTAINIC2 e TCTAFIM2 com os valores encontrados

Dim lErro As Long
Dim sMaior As String, sMenor As String

On Error GoTo Erro_Maior_Menor_Conta

    sMaior = String(STRING_CONTA, 0)
    sMenor = String(STRING_CONTA, 0)

    'se a data in�cio n�o coincide com o in�cio do per�odo inicial
    If CDate(DataInicial.Text) <> CDate(sDtIni_I) Then

        lErro = CF("PlanoConta_Le_Maior_Menor_Conta", sMaior, sMenor)
        If lErro <> SUCESSO Then Error 65337

    End If

    lErro = objRelOpcoes.IncluirParametro("TCTAINIC2", sMenor)
    If lErro <> AD_BOOL_TRUE Then Error 65338

    lErro = objRelOpcoes.IncluirParametro("TCTAFIM2", sMaior)
    If lErro <> AD_BOOL_TRUE Then Error 65339

    Maior_Menor_Conta = SUCESSO

    Exit Function

Erro_Maior_Menor_Conta:

    Maior_Menor_Conta = Err

    Select Case Err

        Case 65337

        Case 65338, 65339

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171984)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCta_I As String, sCta_F As String, sDtIni_I As String, sDtFim_F As String) As Long
'monta a express�o de sele��o

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sCta_I <> "" Then sExpressao = "Conta >= " & Forprint_ConvTexto(sCta_I)

    If sCta_F <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Conta <= " & Forprint_ConvTexto(sCta_F)
    End If

    'se a data inicio n�o coincide com o inicio do per�odo inicial
    If CDate(DataInicial.Text) <> CDate(sDtIni_I) Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LancData >= " & Forprint_ConvData(CDate(DataInicial.Text))
    End If

    'se a data fim n�o coincide com o fim do per�odo final
    If CDate(DataFinal.Text) <> CDate(sDtFim_F) Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LancData <= " & Forprint_ConvData(CDate(DataFinal.Text))
    End If

    Select Case giFilialEmpresa
        Case EMPRESA_TODA
            If giContabGerencial <> 0 Then
                If sExpressao <> "" Then sExpressao = sExpressao & " E "
                sExpressao = sExpressao & "FilialEmpresaLcto < " & Forprint_ConvInt(Abs(giFilialAuxiliar))
            End If
        
        Case Abs(giFilialAuxiliar)
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "FilialEmpresaLcto > " & Forprint_ConvInt(Abs(giFilialAuxiliar))
        
        Case Else
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "FilialEmpresaLcto = " & Forprint_ConvInt(giFilialEmpresa)
    End Select
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = "TipoReg = 1 OU (TipoReg = 2 E " & sExpressao & ")"

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171985)

    End Select

    Exit Function

End Function

Sub Acha_Nome_TSK(sDtIni_I As String)
'acha o nome do arquivo tsk de acordo com Pular Pag e data inicial com in�cio de per�odo

Dim iNumero As Integer

''    iNumero = 1
''
''    If sDtIni_I = DataInicial.Text Then
''
''        If CheckPulaPag.Value = 0 Then iNumero = 2
''
''    Else
''
''        iNumero = 3
''
''        If CheckPulaPag.Value = 0 Then iNumero = 4
''
''    End If
''
''    gobjRelatorio.sNomeTsk = "razao" & CStr(iNumero)
''    gobjRelatorio.sNomeTsk = "razao"
        
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoContaDe = Nothing
    Set objEventoContaAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 65340
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 65341
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 65341
        
        Case 65340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171986)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 65342

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 65343

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as op��es da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 65344
    
        DescCtaInic.Caption = ""
        DescCtaFim.Caption = ""
        CheckPulaPag.Value = 0

    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 65342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 65343, 65344

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171987)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then Error 65345

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 65345

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171988)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then Error 65346

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 65347

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 65348

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 65349
    
    Call BotaoLimpar_Click
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err

        Case 65346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 65347

        Case 65348, 65349
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171989)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 65350
    
    DescCtaInic.Caption = ""
    DescCtaFim.Caption = ""
    CheckPulaPag.Value = 0
    ComboOpcoes.Text = ""

    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 65350
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171990)

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
    If lErro <> SUCESSO Then Error 65351

    Exit Sub

Erro_ContaFinal_Validate:

    Cancel = True


    Select Case Err

        Case 65351

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171991)

    End Select

    Exit Sub

End Sub

Private Sub ContaInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ContaInicial_Validate

    giFocoInicial = 1

    lErro = CF("Conta_Perde_Foco", ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then Error 65352

    Exit Sub

Erro_ContaInicial_Validate:

    Cancel = True


    Select Case Err

        Case 65352

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171992)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 65353

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 65353

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171993)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 65354

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 65354

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171994)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_OpcoesRel_Form_Load

    giFocoInicial = 1

    'inicializa a mascara de conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaInicial)
    If lErro <> SUCESSO Then Error 65355

    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaFinal)
    If lErro <> SUCESSO Then Error 65356
       
'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 65357
    
    Set objEventoContaDe = New AdmEvento
    Set objEventoContaAte = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 65355

        Case 65356

        Case 65357

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171995)

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
        If lErro <> SUCESSO Then Error 65358
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 65358
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171996)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_TvwContas_NodeClick

    sConta = right(Node.Key, Len(Node.Key) - 1)

    lErro = Traz_Conta_Tela(sConta)
    If lErro <> SUCESSO Then Error 65359

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 65359

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171997)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 65360

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 65360
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171998)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 65361

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 65361
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171999)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 65362

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 65362
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172000)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 65363

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 65363
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172001)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_RAZAO
    Set Form_Load_Ocx = Me
    Caption = "Raz�o Aglutinado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRazaoAglut"
    
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



'Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label6, Source, X, Y)
'End Sub
'
'Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
'End Sub

Private Sub DescCtaInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCtaInic, Source, X, Y)
End Sub

Private Sub DescCtaInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCtaInic, Button, Shift, X, Y)
End Sub

'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub

Private Sub DescCtaFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCtaFim, Source, X, Y)
End Sub

Private Sub DescCtaFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCtaFim, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

