VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpRazaoCclOcx 
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   8265
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
      Height          =   1500
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   7995
      Begin MSMask.MaskEdBox ContaInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   330
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaFinal 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Top             =   930
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label DescCtaFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2655
         TabIndex        =   11
         Top             =   930
         Width           =   4965
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.Label DescCtaInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2655
         TabIndex        =   9
         Top             =   330
         Width           =   4950
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
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
      Height          =   1860
      Left            =   120
      TabIndex        =   23
      Top             =   1470
      Width           =   8010
      Begin VB.CheckBox CheckPulaPag 
         Caption         =   "Pula página a cada novo centro de custo"
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
         Left            =   735
         TabIndex        =   32
         Top             =   1470
         Width           =   3915
      End
      Begin MSMask.MaskEdBox CclInicial 
         Height          =   285
         Left            =   720
         TabIndex        =   24
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CclFinal 
         Height          =   285
         Left            =   720
         TabIndex        =   25
         Top             =   975
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label DescCclFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2655
         TabIndex        =   29
         Top             =   975
         Width           =   4935
      End
      Begin VB.Label DescCclInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2655
         TabIndex        =   28
         Top             =   360
         Width           =   4935
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   990
         Width           =   615
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   375
         Width           =   735
      End
   End
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   3585
      Left            =   5955
      TabIndex        =   30
      Top             =   1215
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   6324
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
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
      Height          =   615
      Left            =   4140
      Picture         =   "RelOpRazaoCclOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   135
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRazaoCclOcx.ctx":0102
      Left            =   1290
      List            =   "RelOpRazaoCclOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   255
      Width           =   2610
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6015
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRazaoCclOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRazaoCclOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRazaoCclOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRazaoCclOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   315
      Left            =   1335
      TabIndex        =   14
      Top             =   975
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   315
      Left            =   4110
      TabIndex        =   15
      Top             =   960
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   2535
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   975
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   330
      Left            =   5310
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   3600
      Left            =   5955
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   6350
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label LabelCcls 
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
      Left            =   5955
      TabIndex        =   31
      Top             =   915
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   165
      TabIndex        =   21
      Top             =   1035
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
      Left            =   3045
      TabIndex        =   20
      Top             =   1020
      Width           =   945
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
      Left            =   585
      TabIndex        =   19
      Top             =   315
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
      Left            =   6030
      TabIndex        =   22
      Top             =   915
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "RelOpRazaoCclOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giFocoInicial As Boolean
Dim giFocoInicialCcl As Boolean
Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Private WithEvents objEventoCclDe As AdmEvento
Attribute objEventoCclDe.VB_VarHelpID = -1
Private WithEvents objEventoCclAte As AdmEvento
Attribute objEventoCclAte.VB_VarHelpID = -1
Private WithEvents objEventoContaDe As AdmEvento
Attribute objEventoContaDe.VB_VarHelpID = -1
Private WithEvents objEventoContaAte As AdmEvento
Attribute objEventoContaAte.VB_VarHelpID = -1

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

Function Obtem_Periodo_Exercicio(iPer_I As Integer, iPer_F As Integer, iExercicio As Integer, sDtIni_I As String, sDtFim_F As String) As Long
'a partir das datas ( inicial e final ) encontra o período e o exercício
'as datas devem estar no mesmo exercício
'pega também a data inicial do período inicial e a data final do período final

Dim objPer_I As New ClassPeriodo, objPer_F As New ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_Obtem_Periodo_Exercicio

    'pega o período da Data Inicial
    lErro = CF("Periodo_Le", DataInicial.Text, objPer_I)
    If lErro <> SUCESSO Then Error 13012

    'pega o período da Data Final
    lErro = CF("Periodo_Le", DataFinal.Text, objPer_F)
    If lErro <> SUCESSO Then Error 13013

    'Data Inicial e Final devem estar num mesmo exercício
    If objPer_I.iExercicio <> objPer_F.iExercicio Then Error 13014

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

        Case 13012
            DataInicial.SetFocus

        Case 13013
            DataFinal.SetFocus

        Case 13014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_COM_EXERCICIOS_DIFERENTES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 172077)

    End Select

    Exit Function

End Function

Function Formata_E_Critica_Contas(sCta_I As String, sCta_F As String) As Long
'retorna em sCta_I e sCta_F as contas (inicial e final) formatadas
'Verifica se a conta inicial é maior que a conta final

Dim iCtaPreenchida_I As Integer, iCtaPreenchida_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Contas

    'formata a Conta Inicial
    lErro = CF("Conta_Formata", ContaInicial.Text, sCta_I, iCtaPreenchida_I)
    If lErro <> SUCESSO Then Error 13015
    If iCtaPreenchida_I <> CONTA_PREENCHIDA Then sCta_I = ""

    'formata a Conta Final
    lErro = CF("Conta_Formata", ContaFinal.Text, sCta_F, iCtaPreenchida_F)
    If lErro <> SUCESSO Then Error 13016
    If iCtaPreenchida_F <> CONTA_PREENCHIDA Then sCta_F = ""

    'se ambas as contas estão preenchidas, a conta inicial não pode ser maior que a final
    If iCtaPreenchida_I = CONTA_PREENCHIDA And iCtaPreenchida_F = CONTA_PREENCHIDA Then
        If sCta_I > sCta_F Then Error 13017
    End If

    Formata_E_Critica_Contas = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Contas:

    Formata_E_Critica_Contas = Err

    Select Case Err

        Case 13015
            ContaInicial.SetFocus

        Case 13016
            ContaFinal.SetFocus

        Case 13017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172078)

    End Select

    Exit Function

End Function

Private Function Traz_Conta_Tela(sConta As String) As Long
'verifica e preenche a conta inicial e final com sua descriçao de acordo com o último foco
'sConta deve estar no formato do BD

Dim lErro As Long

On Error GoTo Erro_Traz_Conta_Tela

    If giFocoInicial Then

        lErro = CF("Traz_Conta_MaskEd", sConta, ContaInicial, DescCtaInic)
        If lErro <> SUCESSO Then Error 13030

    Else

        lErro = CF("Traz_Conta_MaskEd", sConta, ContaFinal, DescCtaFim)
        If lErro <> SUCESSO Then Error 13031

    End If

    Traz_Conta_Tela = SUCESSO

    Exit Function

Erro_Traz_Conta_Tela:

    Traz_Conta_Tela = Err

    Select Case Err

        Case 13030
        
        Case 13031

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172079)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpRazao() As Long
'faz a crítica da data inicial e da data final

Dim lErro As Long

On Error GoTo Erro_Critica_Datas_RelOpRazao

    'data inicial não pode ser vazia
    If Len(DataInicial.ClipText) = 0 Then Error 13032

    'data final não pode ser vazia
    If Len(DataFinal.ClipText) = 0 Then Error 13033

    'data inicial não pode ser maior que a data final
    If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 13034

    Critica_Datas_RelOpRazao = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpRazao:

    Critica_Datas_RelOpRazao = Err

    Select Case Err

        Case 13032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)
            DataInicial.SetFocus

        Case 13033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)
            DataFinal.SetFocus

        Case 13034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172080)

    End Select
    
    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arqquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sDescCcl As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 13035

    'pega Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro <> SUCESSO Then Error 13039

    If sParam <> "" Then
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 13039
    End If
    
    CclInicial.PromptInclude = False
    CclInicial.Text = sParam
    CclInicial.PromptInclude = True
    
    DescCclInic.Caption = sDescCcl
    
    'pega Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro <> SUCESSO Then Error 13039

    If sParam <> "" Then
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 13039
    End If
    
    CclFinal.PromptInclude = False
    CclFinal.Text = sParam
    CclFinal.PromptInclude = True
    
    DescCclFim.Caption = sDescCcl

    'pega Conta Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAINIC", sParam)
    If lErro <> SUCESSO Then Error 13036

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then Error 13037

    'pega Conta Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCTAFIM", sParam)
    If lErro <> SUCESSO Then Error 13039

    lErro = CF("Traz_Conta_MaskEd", sParam, ContaFinal, DescCtaFim)
    If lErro <> SUCESSO Then Error 13040

    'pega 'Pula página a cada novo conta' e exibe
    lErro = objRelOpcoes.ObterParametro("TPULAPAGQBR2", sParam)
    If lErro <> SUCESSO Then Error 13041

    If sParam = "S" Then CheckPulaPag.Value = 1
    If sParam = "N" Then CheckPulaPag.Value = 0
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 13045

    DataInicial.PromptInclude = False
    DataInicial.Text = sParam
    DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 13049

    DataFinal.PromptInclude = False
    DataFinal.Text = sParam
    DataFinal.PromptInclude = True

    PreencherParametrosNaTela = SUCESSO
    
    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err
    
    Select Case Err

        Case 13035

        Case 13036, 13039, 13041, 13045, 13049

        Case 13037, 13040

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172081)

    End Select
    
    Exit Function
    
End Function


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bGeraArqTemp As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long, lNumIntRel As Long
Dim iPer_I As Integer, iPer_F As Integer
Dim iExercicio As Integer
Dim sCheck As String
Dim sDtIni_I As String, sDtFim_F As String
Dim sCta_I As String, sCta_F As String
Dim sCcl_I As String, sCcl_F As String
Dim iCclPreenchida_I As Integer, iCclPreenchida_F As Integer

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_RelOpRazao
    If lErro <> SUCESSO Then Error 13050

    lErro = Formata_E_Critica_Contas(sCta_I, sCta_F)
    If lErro <> SUCESSO Then Error 13051

    lErro = Obtem_Periodo_Exercicio(iPer_I, iPer_F, iExercicio, sDtIni_I, sDtFim_F)
    If lErro <> SUCESSO Then Error 13052

    'verifica se o Ccl Inicial é maior que o Ccl Final
    lErro = CF("Ccl_Formata", CclInicial.Text, sCcl_I, iCclPreenchida_I)
    If lErro <> SUCESSO Then Error 13477

    lErro = CF("Ccl_Formata", CclFinal.Text, sCcl_F, iCclPreenchida_F)
    If lErro <> SUCESSO Then Error 13478

    If (iCclPreenchida_I = CCL_PREENCHIDA) And (iCclPreenchida_F = CCL_PREENCHIDA) Then
    
        If sCcl_I > sCcl_F Then Error 13479
    
    End If
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 13056

    'lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(giFilialEmpresa))
    'If lErro <> AD_BOOL_TRUE Then Error 7217

    'lErro = objRelOpcoes.IncluirParametro("TNOMEFILIAL", CStr(gsNomeFilialEmpresa))
    'If lErro <> AD_BOOL_TRUE Then Error 7218

    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then Error 13481

    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then Error 13482

    lErro = objRelOpcoes.IncluirParametro("TCTAINIC", sCta_I)
    If lErro <> AD_BOOL_TRUE Then Error 13057

    lErro = objRelOpcoes.IncluirParametro("TCTAFIM", sCta_F)
    If lErro <> AD_BOOL_TRUE Then Error 13058

    'Pula Página a Cada Nova Conta
    If CheckPulaPag.Value Then
        sCheck = "S"
    Else
        sCheck = "N"
    End If

    lErro = objRelOpcoes.IncluirParametro("TPULAPAGQBR2", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 13059
    
    lErro = objRelOpcoes.IncluirParametro("NPERINIC", CStr(iPer_I))
    If lErro <> AD_BOOL_TRUE Then Error 13060

    lErro = objRelOpcoes.IncluirParametro("NPERFIM", CStr(iPer_F))
    If lErro <> AD_BOOL_TRUE Then Error 13061

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(iExercicio))
    If lErro <> AD_BOOL_TRUE Then Error 13062

    lErro = objRelOpcoes.IncluirParametro("DINICPERINI", sDtIni_I)
    If lErro <> AD_BOOL_TRUE Then Error 13063

    lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13064

    lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13065

    lErro = Maior_Menor_Conta(objRelOpcoes, sDtIni_I)
    If lErro <> SUCESSO Then Error 13066

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCta_I, sCta_F, sDtIni_I, sDtFim_F, sCcl_I, iCclPreenchida_I, sCcl_F, iCclPreenchida_F)
    If lErro <> SUCESSO Then Error 13067

    If bGeraArqTemp Then
    
        'por enquanto nao preciso
'        lErro = CF("RelRazaoCcl_Prepara", giFilialEmpresa, lNumIntRel, sCcl_I, sCcl_F, sCta_I, sCta_F, CDate(DataInicial.Text), CDate(DataFinal.Text))
'        If lErro <> SUCESSO Then Error 54966
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then Error 54964

    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err
    
    Select Case Err

        Case 13050, 13047, 13048, 13481, 13482

        Case 13051

        Case 13052

        Case 13056

        Case 13057, 13058, 13059, 13060, 13061, 13062, 13063, 13064, 13065

        Case 13066

        Case 13067

        Case 13479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172082)
            
    End Select
    
    Exit Function
    
End Function

Function Maior_Menor_Conta(objRelOpcoes As AdmRelOpcoes, sDtIni_I As String) As Long
'acha a menor e a maior conta do BD se a data de início não coincide com o início do período inicial
'preenche TCTAINIC2 e TCTAFIM2 com os valores encontrados

Dim lErro As Long
Dim sMaior As String, sMenor As String

On Error GoTo Erro_Maior_Menor_Conta

    sMaior = String(STRING_CONTA, 0)
    sMenor = String(STRING_CONTA, 0)

    'se a data início não coincide com o início do período inicial
    If CDate(DataInicial.Text) <> CDate(sDtIni_I) Then

        lErro = CF("PlanoConta_Le_Maior_Menor_Conta", sMaior, sMenor)
        If lErro <> SUCESSO Then Error 13068

    End If

    lErro = objRelOpcoes.IncluirParametro("TCTAINIC2", sMenor)
    If lErro <> AD_BOOL_TRUE Then Error 13069

    lErro = objRelOpcoes.IncluirParametro("TCTAFIM2", sMaior)
    If lErro <> AD_BOOL_TRUE Then Error 13070

    Maior_Menor_Conta = SUCESSO

    Exit Function

Erro_Maior_Menor_Conta:

    Maior_Menor_Conta = Err

    Select Case Err

        Case 13068

        Case 13069, 13070

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172083)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCta_I As String, sCta_F As String, sDtIni_I As String, sDtFim_F As String, sCcl_I As String, iCclPreenchida_I As Integer, sCcl_F As String, iCclPreenchida_F As Integer) As Long
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
    
    If sCta_I <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Conta >= " & Forprint_ConvTexto(sCta_I)
    End If
    
    If sCta_F <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Conta <= " & Forprint_ConvTexto(sCta_F)
    End If

    'se a data inicio não coincide com o inicio do período inicial
    If CDate(DataInicial.Text) <> CDate(sDtIni_I) Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LancData >= " & Forprint_ConvData(CDate(DataInicial.Text))
    End If

    'se a data fim não coincide com o fim do período final
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172084)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCclDe = Nothing
    Set objEventoCclAte = Nothing
    Set objEventoContaDe = Nothing
    Set objEventoContaAte = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 24976
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 13085
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 13085
        
        Case 24976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172085)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 13071

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 13072

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47091
    
        DescCclInic.Caption = ""
        DescCclFim.Caption = ""
        DescCtaInic.Caption = ""
        DescCtaFim.Caption = ""
        CheckPulaPag.Value = 0

    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 13071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 13072

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172086)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then Error 13073

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 13073

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172087)

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
    If ComboOpcoes.Text = "" Then Error 13074

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13075

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 13076

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47092
    
    Call BotaoLimpar_Click
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err

        Case 13074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 13075

        Case 13076, 47092
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172088)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47091
    
    DescCclInic.Caption = ""
    DescCclFim.Caption = ""
    DescCtaInic.Caption = ""
    DescCtaFim.Caption = ""
    CheckPulaPag.Value = 0
    ComboOpcoes.Text = ""

    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47091
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172089)

    End Select

    Exit Sub
End Sub

Private Sub CclFinal_GotFocus()
    Mostra_Ccls (True)
End Sub

Private Sub CclInicial_GotFocus()
    Mostra_Ccls (True)
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub ContaFinal_GotFocus()
    Mostra_Ccls (False)
End Sub

Private Sub ContaFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ContaFinal_Validate

    giFocoInicial = 0

    lErro = CF("Conta_Perde_Foco", ContaFinal, DescCtaFim)
    If lErro <> SUCESSO Then Error 13079

    Exit Sub

Erro_ContaFinal_Validate:

    Cancel = True


    Select Case Err

        Case 13079

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172090)

    End Select

    Exit Sub

End Sub

Private Sub ContaInicial_GotFocus()
    Mostra_Ccls (False)
End Sub

Private Sub ContaInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ContaInicial_Validate

    giFocoInicial = 1

    lErro = CF("Conta_Perde_Foco", ContaInicial, DescCtaInic)
    If lErro <> SUCESSO Then Error 13080

    Exit Sub

Erro_ContaInicial_Validate:

    Cancel = True


    Select Case Err

        Case 13080

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172091)

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
        If lErro <> SUCESSO Then Error 13081

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 13081

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172092)

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
        If lErro <> SUCESSO Then Error 13082

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 13082

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172093)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_OpcoesRel_Form_Load

    giFocoInicial = 1
    giFocoInicialCcl = 1
    
    Set objEventoCclDe = New AdmEvento
    Set objEventoCclAte = New AdmEvento
    Set objEventoContaDe = New AdmEvento
    Set objEventoContaAte = New AdmEvento
    
    'inicializa a mascara de centro de custo/lucro inicial
    lErro = Inicializa_Mascara_CclInicial()
    If lErro <> SUCESSO Then Error 54877
    
    'inicializa a mascara de centro de custo/lucro final
    lErro = Inicializa_Mascara_CclFinal()
    If lErro <> SUCESSO Then Error 54878

'    'Inicializa a Lista de Centros de Custo
'    lErro = CF("Carga_Arvore_Ccl", TvwCcls.Nodes)
'    If lErro <> SUCESSO Then Error 13502

    'inicializa a mascara de conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaInicial)
    If lErro <> SUCESSO Then Error 13083

    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaFinal)
    If lErro <> SUCESSO Then Error 13084
       
'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 13086

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 13083, 54877, 54878, 13502

        Case 13084

        Case 13086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172094)

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
        If lErro <> SUCESSO Then Error 40825
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 40825
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 172095)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_TvwContas_NodeClick

    sConta = right(Node.Key, Len(Node.Key) - 1)

    lErro = Traz_Conta_Tela(sConta)
    If lErro <> SUCESSO Then Error 13087

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 13087

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172096)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 13088

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 13088
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172097)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 13089

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 13089
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172098)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 13090

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 13090
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172099)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 13091

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 13091
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172100)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_RAZAO
    Set Form_Load_Ocx = Me
    Caption = "Razão por Centro de Custos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRazaoCcl"
    
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

Private Sub CclFinal_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sCclFormatado As String
Dim iCclPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclFinal_Validate
    
    giFocoInicialCcl = 0
    
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
    
    giFocoInicialCcl = 1
    
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
    
    objCcl.sCcl = right(Node.Key, Len(Node.Key) - 1)
    
    If giFocoInicialCcl = 1 Then
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

Private Sub DescCclFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCclFim, Source, X, Y)
End Sub

Private Sub DescCclFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCclFim, Button, Shift, X, Y)
End Sub

Private Sub DescCclInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCclInic, Source, X, Y)
End Sub

Private Sub DescCclInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCclInic, Button, Shift, X, Y)
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

Private Sub Mostra_Ccls(ByVal bMostrar As Boolean)

'    If bMostrar Then
'
'        LabelCcls.Visible = True
'        TvwCcls.Visible = True
'        TvwContas.Visible = False
'        LabelContas.Visible = False
'
'    Else
'
'        LabelContas.Visible = True
'        TvwContas.Visible = True
'        TvwCcls.Visible = False
'        LabelCcls.Visible = False
'
'    End If
    
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
        
        If Me.ActiveControl Is CclInicial Then Call LabelCclDe_Click
        If Me.ActiveControl Is ContaInicial Then Call LabelContaDe_Click
        If Me.ActiveControl Is CclFinal Then Call LabelCclAte_Click
        If Me.ActiveControl Is ContaFinal Then Call LabelContaAte_Click
    
    End If
    
End Sub
