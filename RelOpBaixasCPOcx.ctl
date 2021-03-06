VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpBaixasCPOcx 
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ScaleHeight     =   4440
   ScaleWidth      =   6645
   Begin VB.CheckBox AgrupaTipoForn 
      Caption         =   "Agrupa por tipo de fornecedor"
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
      TabIndex        =   27
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CheckBox CheckAnalitico 
      Caption         =   "Exibe Devolu��o / Cr�dito"
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
      Left            =   1260
      TabIndex        =   26
      Top             =   3540
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Caption         =   "Conta Corrente"
      Height          =   1545
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   4005
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   1245
         TabIndex        =   25
         Top             =   675
         Width           =   2550
      End
      Begin VB.OptionButton ApenasCta 
         Caption         =   "Apenas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   24
         Top             =   712
         Width           =   1050
      End
      Begin VB.OptionButton TodasCtas 
         Caption         =   "Todas"
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
         Left            =   195
         TabIndex        =   23
         Top             =   285
         Width           =   900
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpBaixasCPOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpBaixasCPOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpBaixasCPOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         HelpContextID   =   1000
         Left            =   1605
         Picture         =   "RelOpBaixasCPOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   10
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
      Left            =   4470
      Picture         =   "RelOpBaixasCPOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   795
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpBaixasCPOcx.ctx":0A96
      Left            =   990
      List            =   "RelOpBaixasCPOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   165
      Width           =   2730
   End
   Begin VB.Frame Frame1 
      Caption         =   "Baixa"
      Height          =   810
      Left            =   120
      TabIndex        =   16
      Top             =   585
      Width           =   3990
      Begin MSComCtl2.UpDown BaixaDeUpDown 
         Height          =   315
         Left            =   1695
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox BaixaDe 
         Height          =   315
         Left            =   540
         TabIndex        =   1
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown BaixaAteUpDown 
         Height          =   315
         Left            =   3615
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox BaixaAte 
         Height          =   315
         Left            =   2460
         TabIndex        =   2
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
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
         Left            =   150
         TabIndex        =   20
         Top             =   345
         Width           =   345
      End
      Begin VB.Label Label2 
         Caption         =   "At�:"
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
         Left            =   2010
         TabIndex        =   19
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Digita��o"
      Height          =   810
      Left            =   120
      TabIndex        =   11
      Top             =   1500
      Width           =   3990
      Begin MSComCtl2.UpDown DigitacaoDeUpDown 
         Height          =   315
         Left            =   1710
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DigitacaoDe 
         Height          =   315
         Left            =   555
         TabIndex        =   3
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown DigitacaoAteUpDown 
         Height          =   315
         Left            =   3600
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DigitacaoAte 
         Height          =   315
         Left            =   2445
         TabIndex        =   4
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
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
         Left            =   150
         TabIndex        =   15
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label4 
         Caption         =   "At�:"
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
         Left            =   1995
         TabIndex        =   14
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   300
      TabIndex        =   21
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "RelOpBaixasCPOcx"
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

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 48587
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 48590
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 48590
        
        Case 48587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167198)

    End Select

    Exit Function

End Function

Private Sub BaixaAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(BaixaAte)

End Sub

Private Sub BaixaDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(BaixaDe)
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 48588
    
    ComboOpcoes.Text = ""
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 48588
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167199)

    End Select

    Exit Sub
   
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro = PreencheComboContas()
    If lErro <> SUCESSO Then Error 59819

    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 48591
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 48591, 59819
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167200)

    End Select

    Exit Sub

End Sub
Function PreencheComboContas() As Long

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeConta As New AdmCodigoNome

On Error GoTo Erro_PreencheComboContas

    'Carrega a Cole��o de Contas
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 59821

    'Preenche a ComboBox CodConta com os objetos da cole��o de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    PreencheComboContas = SUCESSO

    Exit Function
    
Erro_PreencheComboContas:

    PreencheComboContas = Err

    Select Case Err

        Case 59821
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167201)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a op��o de relat�rio com os par�metros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then Error 48593

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 48594

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 48595
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 48596
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 48593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 48594, 48595, 48596
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167202)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 48598

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 48599

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 48598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 48599, 48600, 48601

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167203)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 48602

    'Guarda o nome do tsk a ser executado
    gobjRelatorio.sNomeTsk = "baixapag"
    
    'Se � para agrupar por tipo de fornecedor => altera o nome do tsk a ser executado
    If AgrupaTipoForn.Value = vbChecked Then gobjRelatorio.sNomeTsk = "baixapat"
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 48602

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167204)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sCheckTipo As String
Dim sClienteTipo As String
Dim sCheckCobrador As String
Dim sCobrador As String
Dim sCheckContas As String
Dim sConta As String

On Error GoTo Erro_PreencherRelOp
    
    'Faz a Critica se o Inicial � Maior que o Final, se tudo est� preenchido correto
    lErro = Formata_E_Critica_Parametros(sCheckContas, sConta)
    If lErro <> SUCESSO Then Error 48603

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 48604
             
    'Preenche data baixa inicial
    If Trim(BaixaDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DBXINIC", BaixaDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DBXINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 48605
    
    'Preenche data da baixa Final
    If Trim(BaixaAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DBXFIM", BaixaAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DBXFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 48606
    
    'Preenche a data da digitacao inicial
    If Trim(DigitacaoDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDIGINIC", DigitacaoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDIGINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 48607
    
    'Preenche data da digitacao final
    If Trim(DigitacaoAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDIGFIM", DigitacaoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDIGFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 48608
    
    'Preenche a Conta Corrente
    lErro = objRelOpcoes.IncluirParametro("TCONTA", sConta)
    If lErro <> AD_BOOL_TRUE Then Error 59831
    
    lErro = objRelOpcoes.IncluirParametro("TCONTACORRENTE", ContaCorrente.Text)
    If lErro <> AD_BOOL_TRUE Then Error 59832

    'Preenche com a Opcao Conta Corrente(Todas Contas ou uma Cnta)
    lErro = objRelOpcoes.IncluirParametro("TTODCONTAS", sCheckContas)
    If lErro <> AD_BOOL_TRUE Then Error 59833

    'Preenche com o Exibir Devolu��o / Cr�dito
    lErro = objRelOpcoes.IncluirParametro("NEXIBTIT", CStr(CheckAnalitico.Value))
    If lErro <> AD_BOOL_TRUE Then Error 47822

    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCheckContas, sConta)
    If lErro <> SUCESSO Then Error 48609

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 48603, 48604, 48605, 48606, 48607, 48608, 48609, 59831, 59832, 59833, 47822
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167205)

    End Select

    Exit Function

End Function

Function Formata_E_Critica_Parametros(sCheckContas As String, sConta As String) As Long
'Verifica se os par�metros iniciais s�o maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'Pelo menos um par De/Ate tem que estar Preenchido sen�o -----> Error
    If Trim(BaixaDe.ClipText) = "" Or Trim(BaixaAte.ClipText) = "" Then
        If Trim(DigitacaoDe.ClipText) = "" Or Trim(DigitacaoAte.ClipText) = "" Then Error 48610
    End If
    
    'data da Baixa inicial n�o pode ser maior que a Baixa final
    If Trim(BaixaDe.ClipText) <> "" And Trim(BaixaAte.ClipText) <> "" Then
    
         If CDate(BaixaDe.Text) > CDate(BaixaAte.Text) Then Error 48611
    
    End If
    
    
    'data daDigitacao da Baixa inicial n�o pode ser maior que a data da digitacao da Baixa final
    If Trim(DigitacaoDe.ClipText) <> "" And Trim(DigitacaoAte.ClipText) <> "" Then
    
         If CDate(DigitacaoDe.Text) > CDate(DigitacaoAte.Text) Then Error 48612
    
    End If
    
    'Se a op��o para todas as Contas estiver selecionada
    If TodasCtas.Value = True Then
        sCheckContas = "Todas"
        sConta = ""
    
    'Se a op��o para apenas uma Conta estiver selecionada
    Else
        'Tem que indicar a Conta
        If ContaCorrente.Text = "" Then Error 59838
        sCheckContas = "Uma"
        sConta = ContaCorrente.Text
    
    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
    
        Case 48610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UMA_DATA_NAO_PREENCHIDA", Err)
            
        Case 48611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_INICIAL_MAIOR", Err)
        
        Case 48612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DIGITACAO_BAIXA_INICIAL_MAIOR", Err)
              
        Case 59838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
            ContaCorrente.SetFocus
              
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167206)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCheckContas As String, sConta As String) As Long
'monta a express�o de sele��o de relat�rio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
    
    'Se a op��o para apenas uma conta estiver selecionada
    If sCheckContas = "Uma" Then

        If CheckAnalitico.Value = 1 Then
            sExpressao = "(Conta = " & Forprint_ConvInt(Codigo_Extrai(sConta))
            sExpressao = sExpressao & " OU NumMovCta = 0)"
        Else
            sExpressao = "Conta = " & Forprint_ConvInt(Codigo_Extrai(sConta))
        End If

    End If

    If Trim(BaixaDe.ClipText) <> "" Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Baixa >= " & Forprint_ConvData(CDate(BaixaDe.Text))
    
    End If
    
    If Trim(BaixaAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Baixa <= " & Forprint_ConvData(CDate(BaixaAte.Text))

    End If
        
    If Trim(DigitacaoDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao >= " & Forprint_ConvData(CDate(DigitacaoDe.Text))

    End If
    
    If Trim(DigitacaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao <= " & Forprint_ConvData(CDate(DigitacaoAte.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167207)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String
Dim sCobrador As String
Dim sConta As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 48613
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DBXINIC", sParam)
    If lErro <> SUCESSO Then Error 48614
    
    Call DateParaMasked(BaixaDe, CDate(sParam))
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DBXFIM", sParam)
    If lErro <> SUCESSO Then Error 48615

    Call DateParaMasked(BaixaAte, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDIGINIC", sParam)
    If lErro <> SUCESSO Then Error 48616

    Call DateParaMasked(DigitacaoDe, CDate(sParam))
       
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDIGFIM", sParam)
    If lErro <> SUCESSO Then Error 48617

    Call DateParaMasked(DigitacaoAte, CDate(sParam))
    
    'pega conta e Exibe
    lErro = objRelOpcoes.ObterParametro("TTODCONTAS", sParam)
    If lErro <> SUCESSO Then Error 59841
                   
    If sParam = "Todas" Then
    
        Call TodasCtas_Click
    
    Else
        'se � apenas uma entao exibe esta
        lErro = objRelOpcoes.ObterParametro("TCONTA", sConta)
        If lErro <> SUCESSO Then Error 59842
                            
        ApenasCta.Value = True
        ContaCorrente.Enabled = True
        CheckAnalitico.Enabled = True
        
        If sConta = "" Then
            ContaCorrente.ListIndex = -1
        Else
            ContaCorrente.Text = sConta
        End If
    End If

    lErro = objRelOpcoes.ObterParametro("NEXIBTIT", sParam)
    If lErro <> SUCESSO Then Error 47835
       
    CheckAnalitico.Value = CInt(sParam)

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 48613, 48614, 48615, 48616, 48617, 47835
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167208)

    End Select

    Exit Function

End Function

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'Verifica se foi preenchida a ComboBox
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se est� preenchida com o item selecionado na ComboBox
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub

    'Verifica se existe o �tem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 59846

    'N�o existe o �tem com a STRING na List da ComboBox
    If lErro <> SUCESSO Then Error 59847

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 59846 'Tratado na rotina chamada
    
        Case 59847
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, ContaCorrente.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167209)

    End Select

    Exit Sub

End Sub

Private Sub TodasCtas_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TodasCtas_Click
    
    'Limpa e desabilita a ComboTipo
    ContaCorrente.ListIndex = -1
    ContaCorrente.Text = ""
    ContaCorrente.Enabled = False
    TodasCtas.Value = True
    
    CheckAnalitico.Enabled = False
    
    Exit Sub

Erro_TodasCtas_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167210)

    End Select

    Exit Sub
    
End Sub

Function Define_Padrao() As Long
'preenche padroes (valores default) na tela

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    BaixaDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    BaixaAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    DigitacaoDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DigitacaoAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'defina todas as contas
    Call TodasCtas_Click

    'define Exibir Devolu��o / Cr�dito como Padrao
    CheckAnalitico.Value = 1

    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167211)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ApenasCta_Click()

Dim lErro As Long

On Error GoTo Erro_OptionUmTipo_Click
    
    'Limpa Combo Tipo e Abilita
    ContaCorrente.ListIndex = -1
    ContaCorrente.Enabled = True
    ContaCorrente.SetFocus
    
    CheckAnalitico.Enabled = True
    
    Exit Sub

Erro_OptionUmTipo_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167212)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub BaixaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaixaAte_Validate

    If Len(BaixaAte.ClipText) > 0 Then
        
        lErro = Data_Critica(BaixaAte.Text)
        If lErro <> SUCESSO Then Error 48620

    End If

    Exit Sub

Erro_BaixaAte_Validate:

    Cancel = True


    Select Case Err

        Case 48620

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167213)

    End Select

    Exit Sub

End Sub

Private Sub BaixaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaixaDe_Validate

    If Len(BaixaDe.ClipText) > 0 Then

        lErro = Data_Critica(BaixaDe.Text)
        If lErro <> SUCESSO Then Error 48621

    End If

    Exit Sub

Erro_BaixaDe_Validate:

    Cancel = True


    Select Case Err

        Case 48621

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167214)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DigitacaoAte)

End Sub

Private Sub DigitacaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DigitacaoAte_Validate

    If Len(DigitacaoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(DigitacaoAte.Text)
        If lErro <> SUCESSO Then Error 48622

    End If

    Exit Sub

Erro_DigitacaoAte_Validate:

    Cancel = True


    Select Case Err

        Case 48622

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167215)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DigitacaoDe)

End Sub

Private Sub DigitacaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DigitacaoDe_Validate

    If Len(DigitacaoDe.ClipText) > 0 Then

        lErro = Data_Critica(DigitacaoDe.Text)
        If lErro <> SUCESSO Then Error 48623

    End If

    Exit Sub

Erro_DigitacaoDe_Validate:

    Cancel = True


    Select Case Err

        Case 48623

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167216)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Private Sub DigitacaoDeUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_DigitacaoDeUpDown_DownClick

    lErro = Data_Up_Down_Click(DigitacaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 48624

    Exit Sub

Erro_DigitacaoDeUpDown_DownClick:

    Select Case Err

        Case 48624
            DigitacaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167217)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoDeUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_DigitacaoDeUpDown_UpClick

    lErro = Data_Up_Down_Click(DigitacaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 48625

    Exit Sub

Erro_DigitacaoDeUpDown_UpClick:

    Select Case Err

        Case 48625
            DigitacaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167218)

    End Select

    Exit Sub

End Sub
    
Private Sub BaixaDeUpDoWn_DownClick()

Dim lErro As Long

On Error GoTo Erro_BaixaDeUpDoWn_DownClick

    lErro = Data_Up_Down_Click(BaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 48626

    Exit Sub

Erro_BaixaDeUpDoWn_DownClick:

    Select Case Err

        Case 48626
            BaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167219)

    End Select

    Exit Sub

End Sub

Private Sub BaixaDeUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_BaixaDeUpDown_UpClick

    lErro = Data_Up_Down_Click(BaixaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 48627

    Exit Sub

Erro_BaixaDeUpDown_UpClick:

    Select Case Err

        Case 48627
            BaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167220)

    End Select

    Exit Sub
    
End Sub

Private Sub BaixaAteUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_BaixaAteUpDown_DownClick

    lErro = Data_Up_Down_Click(BaixaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 48628

    Exit Sub

Erro_BaixaAteUpDown_DownClick:

    Select Case Err

        Case 48628
            BaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167221)

    End Select

    Exit Sub

End Sub

Private Sub BaixaAteUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_BaixaAteUpDown_UpClick

    lErro = Data_Up_Down_Click(BaixaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 48629
    
    Exit Sub

Erro_BaixaAteUpDown_UpClick:

    Select Case Err

        Case 48629
            BaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167222)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoAteUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_DigitacaoAteUpDown_DownClick

    lErro = Data_Up_Down_Click(DigitacaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 48630

    Exit Sub

Erro_DigitacaoAteUpDown_DownClick:

    Select Case Err

        Case 48630
            DigitacaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167223)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoAteUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_DigitacaoAteUpDown_UpClick

    lErro = Data_Up_Down_Click(DigitacaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 48631

    Exit Sub

Erro_DigitacaoAteUpDown_UpClick:

    Select Case Err

        Case 48631
            DigitacaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167224)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_BAIXASCP
    Set Form_Load_Ocx = Me
    Caption = "Rela��o de Baixas no Contas a Pagar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBaixasCP"
    
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


Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

