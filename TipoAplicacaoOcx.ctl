VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TipoAplicacaoOcx 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   8295
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5940
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   210
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoAplicacaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoAplicacaoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoAplicacaoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoAplicacaoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListaHistoricoMovConta 
      Height          =   3570
      Left            =   5445
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Frame FrameContabilidade 
      Caption         =   "Contabilidade"
      Height          =   2370
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   5115
      Begin MSMask.MaskEdBox ContaContabilAplicacao 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaReceitaFinanceira 
         Height          =   315
         Left            =   1785
         TabIndex        =   5
         Top             =   1410
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelContaReceitaFinanceira 
         Caption         =   "Conta Receita:"
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
         Left            =   405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   1425
         Width           =   1335
      End
      Begin VB.Label LabelContaContabilAplicacao 
         AutoSize        =   -1  'True
         Caption         =   "Conta Aplicação:"
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label DescContaContabil 
         Height          =   345
         Left            =   1755
         TabIndex        =   19
         Top             =   825
         Width           =   3015
      End
      Begin VB.Label LabelDescContaContabil 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Height          =   225
         Left            =   720
         TabIndex        =   20
         Top             =   825
         Width           =   945
      End
      Begin VB.Label DescContaReceita 
         Height          =   330
         Left            =   1710
         TabIndex        =   21
         Top             =   1845
         Width           =   3015
      End
      Begin VB.Label LabelDescContaReceita 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Height          =   210
         Left            =   630
         TabIndex        =   22
         Top             =   1875
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   1470
      Left            =   165
      TabIndex        =   15
      Top             =   120
      Width           =   5100
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1905
         Picture         =   "TipoAplicacaoOcx.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   345
         Width           =   300
      End
      Begin VB.CheckBox Inativo 
         Caption         =   "Inativo"
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
         Left            =   2970
         TabIndex        =   2
         Top             =   390
         Width           =   975
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   300
         Left            =   1425
         TabIndex        =   3
         Top             =   930
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Top             =   330
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   420
         TabIndex        =   23
         Top             =   960
         Width           =   930
      End
      Begin VB.Label LabelCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   690
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.TextBox HistoricoPadrao 
      Height          =   285
      Left            =   2220
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4500
      Width           =   2355
   End
   Begin VB.ListBox ListaTipoAplicacao 
      Height          =   3570
      ItemData        =   "TipoAplicacaoOcx.ctx":0A7E
      Left            =   5475
      List            =   "TipoAplicacaoOcx.ctx":0A80
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1170
      Width           =   2610
   End
   Begin MSComctlLib.TreeView TvwConta 
      Height          =   3570
      Left            =   5460
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   6297
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label LabelTiposDeAplicacao 
      AutoSize        =   -1  'True
      Caption         =   "Tipos de Aplicação"
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
      Left            =   5505
      TabIndex        =   25
      Top             =   945
      Width           =   1650
   End
   Begin VB.Label LabelHistoricos 
      AutoSize        =   -1  'True
      Caption         =   "Históricos"
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
      Left            =   5565
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LabelPlanoDeContas 
      AutoSize        =   -1  'True
      Caption         =   "Plano de Conta"
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
      Left            =   5520
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label5 
      Caption         =   "Histórico no Extrato:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   28
      Top             =   4530
      Width           =   1815
   End
End
Attribute VB_Name = "TipoAplicacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoContaContabilAplicacao As AdmEvento
Attribute objEventoContaContabilAplicacao.VB_VarHelpID = -1
Private WithEvents objEventoContaReceitaFinanceira As AdmEvento
Attribute objEventoContaReceitaFinanceira.VB_VarHelpID = -1
Private WithEvents objEventoTipoAplicacao As AdmEvento
Attribute objEventoTipoAplicacao.VB_VarHelpID = -1

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Código do Tipo de aplicação está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 15186
    
    objTiposDeAplicacao.iCodigo = CInt(Codigo.Text)
    
    'Envia mensagem pedindo confirmação de exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPOAPLICACAO")
    
    If vbMsgRes = vbYes Then
        
        'Exclui o Tipo de aplicação da tabela
        lErro = CF("TiposDeAplicacao_Exclui", objTiposDeAplicacao)
        If lErro <> SUCESSO Then Error 15187
                
        'Exclui o Tipo de aplicação da ListBox
        Call ListaTiposDeAplicacao_Exclui(objTiposDeAplicacao.iCodigo)
        
        'Limpa a tela
        Call Limpa_Tela_TipoAplicacao
        
        iAlterado = 0
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 15186
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
                   
        Case 15187
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174687)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoGravar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Gravao Tipo de aplicação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 15204
    
    'Limpa a tela
    Call Limpa_Tela_TipoAplicacao
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:
    
    Select Case Err
    
        Case 15204
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174688)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma se deseja salvar alterações
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 15201

    'Limpa a tela
    Call Limpa_Tela_TipoAplicacao
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 15201
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174689)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático para o código de Tipo de Aplicação
    lErro = TipoAplicacao_Automatico(iCodigo)
    If lErro <> SUCESSO Then Error 57750

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57750
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174690)
    
    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    'Exibe a ListBox de Tipos de aplicação
    ListaTipoAplicacao.Visible = True
    ListaHistoricoMovConta.Visible = False
    TvwConta.Visible = False
    LabelTiposDeAplicacao.Visible = True
    LabelHistoricos.Visible = False
    LabelPlanoDeContas.Visible = False

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(Codigo.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(Codigo.Text) Then Error 55964

        'Verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then Error 55965

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 55964, 55965
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174691)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabilAplicacao_GotFocus()
    
    'Exibe a árvore de Plano de Contas
    TvwConta.Visible = True
    ListaTipoAplicacao.Visible = False
    ListaHistoricoMovConta.Visible = False
    LabelPlanoDeContas.Visible = True
    LabelHistoricos.Visible = False
    LabelTiposDeAplicacao.Visible = False
    TvwConta.Tag = CONTA_APLICACAO

End Sub

Private Sub ContaContabilAplicacao_Validate(Cancel As Boolean)
'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    
Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_ContaContabilAplicacao_Validate

    'Informa que o último campo clickado foi ContaContabilAplicacao
    TvwConta.Tag = CONTA_APLICACAO
       
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabilAplicacao.Text, ContaContabilAplicacao.ClipText, objPlanoConta, MODULO_TESOURARIA)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 39803
        
    If lErro = SUCESSO Then
        
        sContaFormatada = objPlanoConta.sConta
            
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
            
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 39804
            
        ContaContabilAplicacao.PromptInclude = False
        ContaContabilAplicacao.Text = sContaMascarada
        ContaContabilAplicacao.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
       
        'Critica o formato da Conta
        lErro = CF("Conta_Critica", ContaContabilAplicacao.Text, sContaFormatada, objPlanoConta, MODULO_TESOURARIA)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 15105
                        
        'Critica se a Conta está cadastrada
        If lErro = 5700 Then Error 15321
            
    End If
    
    'Coloca a descrição da Conta no Label
    DescContaContabil.Caption = objPlanoConta.sDescConta
    
    Exit Sub

Erro_ContaContabilAplicacao_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 15105, 39803

        Case 15321
            'Confirma se deseja cadastrar a conta
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilAplicacao.Text)
                
            If vbMsgRes = vbYes Then
                
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("PlanoConta", objPlanoConta)
            
            End If
        
        Case 39804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174692)

    End Select
    
    Exit Sub
        
End Sub

Private Sub ContaReceitaFinanceira_GotFocus()
    
    'Exibe a árvore de Plano de Contas
    TvwConta.Visible = True
    ListaTipoAplicacao.Visible = False
    ListaHistoricoMovConta.Visible = False
    LabelPlanoDeContas.Visible = True
    LabelHistoricos.Visible = False
    LabelTiposDeAplicacao.Visible = False
    TvwConta.Tag = CONTA_RECEITA

End Sub

Private Sub ContaReceitaFinanceira_Validate(Cancel As Boolean)
'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    
Dim lErro As Long
Dim sContaMascarada As String
Dim sContaFormatada As String
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_ContaReceitaFinanceira_Validate
    
    'Informa que o último campo clickado foi ContaReceitaFinanceira
    TvwConta.Tag = CONTA_RECEITA
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaReceitaFinanceira.Text, ContaReceitaFinanceira.ClipText, objPlanoConta, MODULO_TESOURARIA)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 39805
        
    If lErro = SUCESSO Then
        
        sContaFormatada = objPlanoConta.sConta
            
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
            
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 39806
            
        ContaReceitaFinanceira.PromptInclude = False
        ContaReceitaFinanceira.Text = sContaMascarada
        ContaReceitaFinanceira.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
    
        'Critica o formato da Conta
        lErro = CF("Conta_Critica", ContaReceitaFinanceira.Text, sContaFormatada, objPlanoConta, MODULO_TESOURARIA)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 15106
        
        'Critica se a Conta está cadastrada
        If lErro = 5700 Then Error 15322
        
    End If
    
    'Coloca a descrição da Conta no Label
    DescContaReceita.Caption = objPlanoConta.sDescConta
    
    Exit Sub

Erro_ContaReceitaFinanceira_Validate:

    Cancel = True

    Select Case Err

        Case 15106, 39805
        
        Case 15322
            'Confirma se deseja cadastrar a conta
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaReceitaFinanceira.Text)
                
            If vbMsgRes = vbYes Then
                
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("PlanoConta", objPlanoConta)
            
            End If
            
        Case 39806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174693)

    End Select
        
    Exit Sub

End Sub

Private Sub Descricao_GotFocus()
    
    'Exibe a ListBox de Tipos de aplicação
    ListaTipoAplicacao.Visible = True
    ListaHistoricoMovConta.Visible = False
    TvwConta.Visible = False
    LabelTiposDeAplicacao.Visible = True
    LabelHistoricos.Visible = False
    LabelPlanoDeContas.Visible = False

End Sub

Public Sub Form_Activate()

    'Carrega os índices da Tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

'Sai da tela, mas antes verifica se houveram alterações e pede confirmação se deseja salvar

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objEventoTipoAplicacao = Nothing
    Set objEventoContaContabilAplicacao = Nothing
    Set objEventoContaReceitaFinanceira = Nothing
    
     'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Private Sub HistoricoPadrao_GotFocus()
    
    'Exibe a ListBox de Históricos Padrão de Movimentação de Conta
    ListaHistoricoMovConta.Visible = True
    TvwConta.Visible = False
    ListaTipoAplicacao.Visible = False
    LabelHistoricos.Visible = True
    LabelTiposDeAplicacao.Visible = False
    LabelPlanoDeContas.Visible = False

End Sub

Private Sub Inativo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabilAplicacao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaReceitaFinanceira_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodNome As AdmCodigoNome
Dim objTiposDeAplicacao As ClassTiposDeAplicacao
Dim colTiposDeAplicacao As New Collection
Dim sListBoxItem As String
Dim sEspacos As String

On Error GoTo Erro_Form_Load
        
    Set objEventoTipoAplicacao = New AdmEvento
    Set objEventoContaContabilAplicacao = New AdmEvento
    Set objEventoContaReceitaFinanceira = New AdmEvento
                
    'Carrega a ListBox de Tipos de Aplicação
    lErro = CF("TiposDeAplicacao_Le_Todos", colTiposDeAplicacao)
    If lErro <> SUCESSO Then Error 15146
    
    For Each objTiposDeAplicacao In colTiposDeAplicacao
        
        sEspacos = Space(STRING_CODIGO_TIPOAPLICACAO - Len(CStr(objTiposDeAplicacao.iCodigo)))
        ListaTipoAplicacao.AddItem (sEspacos & CStr(objTiposDeAplicacao.iCodigo) & SEPARADOR & objTiposDeAplicacao.sDescricao)
        ListaTipoAplicacao.ItemData(ListaTipoAplicacao.NewIndex) = objTiposDeAplicacao.iCodigo
    
    Next
        
    'Carrega o frame de Contabilidade e a árvore de Contas
    lErro = Carrega_Contabilidade_TipoAplicacao()
    If lErro <> SUCESSO Then Error 15145

    'Carrega a ListBox de Históricos Padrão de Movimentação de Conta
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_HISTORICO, colCodigoNome)
    If lErro <> SUCESSO Then Error 15147

    'Preenche a ListBox com Históricos existentes na coleção
    For Each objCodNome In colCodigoNome

        'Espaços que faltam para completar tamanho STRING_CODIGO_HISTORICO
        sListBoxItem = Space(STRING_CODIGO_HISTORICO - Len(CStr(objCodNome.iCodigo)))

        'Concatena Código e Descrição do Histórico
        sListBoxItem = sListBoxItem & CStr(objCodNome.iCodigo)
        sListBoxItem = sListBoxItem & SEPARADOR & Trim(objCodNome.sNome)

        ListaHistoricoMovConta.AddItem sListBoxItem

    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case 15145, 15146, 15147
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174694)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objTiposDeAplicacao As ClassTiposDeAplicacao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se objTiposDeAplicacao estiver preenchido
    If Not (objTiposDeAplicacao Is Nothing) Then
       
        'Carrega os dados da memória para a tela
        lErro = Traz_Dados_Tela(objTiposDeAplicacao)
        If lErro <> SUCESSO And lErro <> 15327 Then Error 15148
        
        If lErro = 15327 Then Error 15329
                
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 15148
        
        Case 15329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_INEXISTENTE", Err, objTiposDeAplicacao.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174695)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Sub HistoricoPadrao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelCodigo_Click()
'Exibe Browse de Tipos de aplicação

Dim objTiposDeAplicacao As New ClassTiposDeAplicacao
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_LabelCodigo_Click

    lErro = Move_Tela_Memoria(objTiposDeAplicacao)
    If lErro <> SUCESSO Then Error 15323
        
    'Chama a tela com a lista de Tipos de aplicação
    Call Chama_Tela("TipoAplicacaoLista", colSelecao, objTiposDeAplicacao, objEventoTipoAplicacao)

    Exit Sub
    
Erro_LabelCodigo_Click:

    Select Case Err
    
        Case 15323
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174696)
        
    End Select

    Exit Sub
        
End Sub

Private Sub LabelContaContabilAplicacao_Click()
'Chama browse do plano de contas

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_LabelContaContabilAplicacao_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabilAplicacao.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 57752

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaTESLista", colSelecao, objPlanoConta, objEventoContaContabilAplicacao)

    Exit Sub

Erro_LabelContaContabilAplicacao_Click:

    Select Case Err

        Case 57752
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174697)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabilAplicacao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabilAplicacao_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        ContaContabilAplicacao.Text = ""
    Else
        ContaContabilAplicacao.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 57754

        ContaContabilAplicacao.Text = sContaEnxuta

        ContaContabilAplicacao.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabilAplicacao_evSelecao:

    Select Case Err

        Case 57754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174698)

    End Select

    Exit Sub

End Sub

Private Sub LabelContaReceitaFinanceira_Click()
'Chama browse do plano de contas

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_LabelContaReceitaFinanceira_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaReceitaFinanceira.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 57753

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaTESLista", colSelecao, objPlanoConta, objEventoContaReceitaFinanceira)

    Exit Sub

Erro_LabelContaReceitaFinanceira_Click:

    Select Case Err

        Case 57753
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174699)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaReceitaFinanceira_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaReceitaFinanceira_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        ContaReceitaFinanceira.Text = ""
    Else
        ContaReceitaFinanceira.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 57755

        ContaReceitaFinanceira.Text = sContaEnxuta

        ContaReceitaFinanceira.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaReceitaFinanceira_evSelecao:

    Select Case Err

        Case 57755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174700)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub ListaHistoricoMovConta_KeyPress(KeyAscii As Integer)
'Passa o código do Histórico Padrão para a tela.

    'Se não houver histórico selecionado na Listbox de Históricos Padrão de Movimentação de Conta
    If ListaHistoricoMovConta.ListIndex = -1 Then Exit Sub
    
    'Se a tecla pressionada for Enter
    If KeyAscii = ENTER_KEY Then
    
        'Executa o mesmo procedimento que o duplo click
        Call ListaHistoricoMovConta_DblClick
    
    End If

End Sub

Private Sub ListaTipoAplicacao_KeyPress(KeyAscii As Integer)
'Passa os dados do Tipo de aplicação para a tela
    
    'Se não houver Tipo de aplicação selecionado na Listbox de Tipos de aplicação
    If ListaTipoAplicacao.ListIndex = -1 Then Exit Sub
    
    'Se a tecla pressionada for Enter
    If KeyAscii = ENTER_KEY Then
    
        'Executa o mesmo procedimento que o duplo click
        Call ListaTipoAplicacao_DblClick
    
    End If

End Sub

Private Sub objEventoTipoAplicacao_evSelecao(obj1 As Object)
'Evento referente ao Browse de Tipos de aplicação exibido no duplo click do Label Código
    
Dim objTiposDeAplicacao As ClassTiposDeAplicacao
Dim lErro As Long

On Error GoTo Erro_objEventoTipoAplicacao_evSelecao

    Set objTiposDeAplicacao = obj1
    
    'Coloca na tela os dados do Tipo de aplicação passado pelo Obj
    lErro = Traz_Dados_Tela(objTiposDeAplicacao)
    If lErro <> SUCESSO And lErro <> 15327 Then Error 15212
    
    iAlterado = 0
        
    Me.Show
    
    Exit Sub
    
Erro_objEventoTipoAplicacao_evSelecao:

    Select Case Err
    
        Case 15212
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174701)
        
    End Select

    Exit Sub
        
End Sub

Private Sub ListaHistoricoMovConta_DblClick()
'Passa o código do Histórico Padrão para a tela
 
Dim sListBoxItem As String
Dim lPosicaoSeparador As Long

    'Se não houver histórico selecionado na Listbox de Históricos Padrão de Movimentação de Conta
    If ListaHistoricoMovConta.ListIndex = -1 Then Exit Sub
    
    'Acha a posição do separador (-)
    lPosicaoSeparador = InStr(ListaHistoricoMovConta.Text, SEPARADOR)
    
    'Coloca o Histórico na Tela
    HistoricoPadrao.Text = Mid(ListaHistoricoMovConta.Text, lPosicaoSeparador + 1)
    
    Exit Sub
    
End Sub

Private Sub ListaTipoAplicacao_DblClick()
'Passa os dados do Tipo de aplicação para a tela

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao
Dim sContaEnxutaAplic As String
Dim sContaEnxutaRec As String

On Error GoTo Erro_ListaTipoAplicacao_DblClick
    
    'Inicializa contas
    sContaEnxutaAplic = String(STRING_CONTA, 0)
    sContaEnxutaRec = String(STRING_CONTA, 0)
    
    'Se não houver Tipo de aplicação selecionado na Listbox de Tipos de aplicação
    If ListaTipoAplicacao.ListIndex = -1 Then Exit Sub
        
    objTiposDeAplicacao.iCodigo = ListaTipoAplicacao.ItemData(ListaTipoAplicacao.ListIndex)
            
    'Verifica se o Tipo de aplicação existe
    lErro = CF("TiposDeAplicacao_Le", objTiposDeAplicacao)
    If lErro <> SUCESSO And lErro <> 15068 Then Error 15331
    
    'Se Tipo de aplicação não existe
    If lErro = 15068 Then Error 15327
    
    lErro = Traz_Dados_Tela(objTiposDeAplicacao)
    If lErro <> SUCESSO And lErro <> 15327 Then Error 15328
       
    If lErro = 15327 Then Error 15348
    
    'Fecha o comando das setas, se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
      
    iAlterado = 0
    
    Exit Sub
    
Erro_ListaTipoAplicacao_DblClick:

    Select Case Err
        
        Case 15327, 15328, 15331
        
        Case 15348
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_INEXISTENTE", Err, ListaTipoAplicacao.ListIndex)
            
            'Exclui o ítem da ListBox
            ListaTipoAplicacao.RemoveItem (ListaTipoAplicacao.ListIndex)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174702)

    End Select
    
    Exit Sub
    
End Sub

Private Sub TvwConta_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwConta_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta_Modulo1", objNode, TvwConta.Nodes, MODULO_TESOURARIA)
        If lErro <> SUCESSO Then Error 40809
        
    End If
    
    Exit Sub
    
Erro_TvwConta_Expand:

    Select Case Err
    
        Case 40809
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174703)
        
    End Select
        
    Exit Sub
    
End Sub


Private Sub TvwConta_NodeClick(ByVal Node As MSComctlLib.Node)
'Passa os dados da Conta para a tela

Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_TvwConta_NodeClick

    sCaracterInicial = left(Node.Key, 1)

    If sCaracterInicial = CONTA_ANALITICA_ABREV Then

        sConta = right(Node.Key, Len(Node.Key) - 1)

        sContaEnxuta = String(STRING_CONTA, 0)
        
        lPosicaoSeparador = InStr(Node.Text, SEPARADOR)

        'Retorna a Conta em formato enxuto (sem pontos e vírgulas)
        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 15185
            
        'Se o último campo clickado foi ContaAplicacao
        If TvwConta.Tag = CONTA_APLICACAO Then
        
            'Carrega os dados da Conta no campo ContaContabilAplicacao
            ContaContabilAplicacao.PromptInclude = False
            ContaContabilAplicacao.Text = sContaEnxuta
            ContaContabilAplicacao.PromptInclude = True
            DescContaContabil.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
        
        'Se o último campo clicado foi ContaReceita
        ElseIf TvwConta.Tag = CONTA_RECEITA Then

            'Carrega os dados da Conta no campo ContaReceitaFinanceira
            ContaReceitaFinanceira.PromptInclude = False
            ContaReceitaFinanceira.Text = sContaEnxuta
            ContaReceitaFinanceira.PromptInclude = True
            DescContaReceita.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
        
        End If
        
    End If

    Exit Sub
           
Erro_TvwConta_NodeClick:

    Select Case Err
                
        Case 15185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174704)
            
    End Select
        
    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava o Tipo de aplicação na tabela TiposDeAplicacao

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 15116
    
    'Verifica se a Descrição está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then Error 15117
    
    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objTiposDeAplicacao)
    If lErro <> SUCESSO Then Error 15169
    
    'Grava o Tipo de aplicação
    lErro = CF("TiposDeAplicacao_Grava", objTiposDeAplicacao)
    If lErro <> SUCESSO Then Error 15123
    
    'Fecha o comando das setas, se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'Exclui o Tipo de aplicação da ListBox Tipos de aplicação, se já estiver lá
    Call ListaTiposDeAplicacao_Exclui(objTiposDeAplicacao.iCodigo)
    
    'Adiciona o Tipo de aplicação na ListBox Tipos de aplicação
    Call ListaTiposDeAplicacao_Adiciona(objTiposDeAplicacao)
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 15116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
            
        Case 15117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)
            
        Case 15123, 15169
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174705)

     End Select
        
     Exit Function

End Function

Private Function Limpa_Tela_TipoAplicacao()
'Limpa todos os campos da tela TipoAplicacao
        
    Call Limpa_Tela(Me)
    
    Codigo.Text = ""
    
    'Limpa os campos restantes
    Inativo.Value = 0
    DescContaContabil.Caption = ""
    DescContaReceita.Caption = ""

End Function

Private Sub ListaTiposDeAplicacao_Exclui(iCodigo As Integer)
'Exclui ítem da ListBox Tipos de aplicação

Dim iIndice As Integer

    'Percorre todos os itens da ListBox Tipos de aplicação
    For iIndice = 0 To ListaTipoAplicacao.ListCount - 1
    
        'Se o ItemData do ítem for igual ao Código do Tipo de aplicação em questão
        If ListaTipoAplicacao.ItemData(iIndice) = iCodigo Then
        
            'Remove o ítem da List Box Tipos de aplicação
            ListaTipoAplicacao.RemoveItem (iIndice)
            Exit For
        
        End If
    
    Next

End Sub

Private Sub ListaTiposDeAplicacao_Adiciona(objTiposDeAplicacao As ClassTiposDeAplicacao)
'Adiciona ítem na ListBox Tipos de aplicação
    
Dim sEspacos As String

    sEspacos = Space(STRING_CODIGO_TIPOAPLICACAO - Len(CStr(objTiposDeAplicacao.iCodigo)))
    
    'Adiciona a Descrição do Tipo de aplicação na ListBox Tipos de aplicação
    ListaTipoAplicacao.AddItem (sEspacos & CStr(objTiposDeAplicacao.iCodigo) & SEPARADOR & objTiposDeAplicacao.sDescricao)
    
    'Coloca o Código do Tipo de aplicação no ItemData do ítem adicionado
    ListaTipoAplicacao.ItemData(ListaTipoAplicacao.NewIndex) = objTiposDeAplicacao.iCodigo
            
End Sub

Function Carrega_Contabilidade_TipoAplicacao() As Long
'Carrega na tela os dados referentes à Contabilidade.

Dim lErro As Long
Dim sMascaraContas As String

On Error GoTo Erro_Carrega_Contabilidade_TipoAplicacao

    'Se o módulo Contabilidade estiver ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        'Carrega a árvore de Contas Contábeis
        lErro = CF("Carga_Arvore_Conta_Modulo", TvwConta.Nodes, MODULO_TESOURARIA)
        If lErro <> SUCESSO Then Error 15325
    
        'Inicializa a máscara das Contas
        sMascaraContas = String(STRING_CONTA, 0)
        
        'Lê a máscara das Contas
        lErro = MascaraConta(sMascaraContas)
        If lErro <> SUCESSO Then Error 15090
    
        'Atribui a máscara aos campos Conta Aplicação e Conta Receita
        ContaContabilAplicacao.Mask = sMascaraContas
        ContaReceitaFinanceira.Mask = sMascaraContas
    
    Else
        
       'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraContas)
        If lErro <> SUCESSO Then Error 15090
    
        'Atribui a máscara aos campos Conta Aplicação e Conta Receita
        ContaContabilAplicacao.Mask = sMascaraContas
        ContaReceitaFinanceira.Mask = sMascaraContas
        
        'Desabilita o frame de Contabilidade
        FrameContabilidade.Enabled = False
        ContaContabilAplicacao.Enabled = False
        ContaReceitaFinanceira.Enabled = False
        LabelContaContabilAplicacao.Enabled = False
        LabelContaReceitaFinanceira.Enabled = False
        LabelDescContaContabil.Enabled = False
        LabelDescContaReceita.Enabled = False
    End If
    
    Carrega_Contabilidade_TipoAplicacao = SUCESSO
    
    Exit Function
    
Erro_Carrega_Contabilidade_TipoAplicacao:

    Carrega_Contabilidade_TipoAplicacao = Err
    
    Select Case Err
    
        Case 15090, 15325
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174706)
            
    End Select
    
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada
    sTabela = "TiposDeAplicacao"
    
    If Len(Trim(Codigo.ClipText)) > 0 Then
        objTiposDeAplicacao.iCodigo = CInt(Codigo.Text)
    Else
        objTiposDeAplicacao.iCodigo = 0
    End If
    
    If Len(Descricao.Text) > 0 Then
        objTiposDeAplicacao.sDescricao = Descricao.Text
    Else
        objTiposDeAplicacao.sDescricao = String(STRING_TIPOAPLIC_DESCRICAO, 0)
    End If
        
    If Inativo.Value = TIPOAPLICACAO_INATIVO Then
        objTiposDeAplicacao.iInativo = TIPOAPLICACAO_INATIVO
    Else
        objTiposDeAplicacao.iInativo = TIPOAPLICACAO_ATIVO
    End If
    
    If Len(HistoricoPadrao.Text) > 0 Then
        objTiposDeAplicacao.sHistorico = HistoricoPadrao.Text
    Else
        objTiposDeAplicacao.sHistorico = String(STRING_TIPOAPLIC_HISTORICO, 0)
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    
    colCampoValor.Add "Codigo", objTiposDeAplicacao.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objTiposDeAplicacao.sDescricao, STRING_TIPOAPLIC_DESCRICAO, "Descricao)"
    colCampoValor.Add "Inativo", objTiposDeAplicacao.iInativo, 0, "Inativo"
    colCampoValor.Add "Historico", objTiposDeAplicacao.sHistorico, STRING_TIPOAPLIC_HISTORICO, "Historico"
    colCampoValor.Add "ContaContabilAplicacao", objTiposDeAplicacao.sContaAplicacao, STRING_CONTA, "ContaContabilAplicacao"
    colCampoValor.Add "ContaReceitaFinanceira", objTiposDeAplicacao.sContaReceita, STRING_CONTA, "ContaReceitaFinanceira"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174707)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao

On Error GoTo Erro_Tela_Preenche

    objTiposDeAplicacao.iCodigo = colCampoValor.Item("Codigo").vValor
    
    If objTiposDeAplicacao.iCodigo <> 0 Then
    
        objTiposDeAplicacao.sDescricao = colCampoValor.Item("Descricao").vValor
        objTiposDeAplicacao.iInativo = colCampoValor.Item("Inativo").vValor
        objTiposDeAplicacao.sHistorico = colCampoValor.Item("Historico").vValor
        objTiposDeAplicacao.sContaAplicacao = colCampoValor.Item("ContaContabilAplicacao").vValor
        objTiposDeAplicacao.sContaReceita = colCampoValor.Item("ContaReceitaFinanceira").vValor
        
        'Preenche a tela com os dados retornados
        lErro = Traz_Dados_Tela(objTiposDeAplicacao)
        If lErro <> SUCESSO Then Error 34672
                
        iAlterado = 0
        
    End If

    Exit Sub
    
Erro_Tela_Preenche:

    Select Case Err
        
        Case 34672
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174708)
    
    End Select
    
    Exit Sub

End Sub

Private Function Traz_Dados_Tela(objTiposDeAplicacao As ClassTiposDeAplicacao) As Long
'Traz para a tela os dados do Tipo de aplicação passado como parâmetro

Dim lErro As Long

On Error GoTo Erro_Traz_Dados_Tela

    Call Limpa_Tela_TipoAplicacao

    'Carrega a tela com os dados do Tipo de aplicação contidos no Obj
    Codigo.Text = CStr(objTiposDeAplicacao.iCodigo)
    Descricao.Text = objTiposDeAplicacao.sDescricao
    HistoricoPadrao.Text = objTiposDeAplicacao.sHistorico
    Inativo.Value = objTiposDeAplicacao.iInativo
    
    'Traz para a tela os dados referentes à Contabilidade
    lErro = Traz_Dados_Tela_Contabilidade(objTiposDeAplicacao)
    If lErro <> SUCESSO Then Error 15152
    
    Traz_Dados_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Dados_Tela:

    Traz_Dados_Tela = Err
    
    Select Case Err
    
        Case 15152
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174709)
    
    End Select
            
    Exit Function
            
End Function

Private Function Traz_Dados_Tela_Contabilidade(objTiposDeAplicacao As ClassTiposDeAplicacao)
'Traz para a tela os dados do Tipo de aplicação referentes à Contabilidade
    
Dim lErro As Long
Dim sContaEnxutaAplic As String
Dim sContaEnxutaRec As String
Dim sDescContaAplicacao As String
Dim sDescContaReceita As String
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Traz_Dados_Tela_Contabilidade
    
    'Se o módulo de Contabilidade estiver ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        'Inicializa as Contas
        sContaEnxutaAplic = String(STRING_CONTA, 0)
        sContaEnxutaRec = String(STRING_CONTA, 0)
        sDescContaAplicacao = String(STRING_CONTA_DESCRICAO, 0)
        sDescContaReceita = String(STRING_CONTA_DESCRICAO, 0)
        
        If Len(Trim(objTiposDeAplicacao.sContaAplicacao)) <> 0 Then
         
            'Formata ContaContabilAplicacao
            lErro = Mascara_RetornaContaEnxuta(objTiposDeAplicacao.sContaAplicacao, sContaEnxutaAplic)
            If lErro <> SUCESSO Then Error 15149
                    
            'Lê os dados da Conta passada como parâmetro
            lErro = CF("PlanoConta_Le_Conta1", objTiposDeAplicacao.sContaAplicacao, objPlanoConta)
            If lErro <> SUCESSO And lErro <> 11807 Then Error 15345
        
            'Verifica se a Conta existe
            If lErro = 11807 Then Error 15346
                                   
        Else
            
            sContaEnxutaAplic = ""
            objPlanoConta.sDescConta = ""
            
        End If
        
        'Carrega ContaContabilAplicacao na tela
        ContaContabilAplicacao.PromptInclude = False
        ContaContabilAplicacao.Text = sContaEnxutaAplic
        ContaContabilAplicacao.PromptInclude = True
        
        'Carrega a Descrição na tela
        DescContaContabil.Caption = objPlanoConta.sDescConta
         
        If Len(Trim(objTiposDeAplicacao.sContaReceita)) <> 0 Then
         
            'Formata ContaReceitaFinanceira
            lErro = Mascara_RetornaContaEnxuta(objTiposDeAplicacao.sContaReceita, sContaEnxutaRec)
            If lErro <> SUCESSO Then Error 15150
            
            'Lê os dados da Conta passada como parâmetro
            lErro = CF("PlanoConta_Le_Conta1", objTiposDeAplicacao.sContaReceita, objPlanoConta)
            If lErro <> SUCESSO And lErro <> 11807 Then Error 15343
        
            'Verifica se a Conta existe
            If lErro = 11807 Then Error 15344
                      
        Else
            
            sContaEnxutaRec = ""
            objPlanoConta.sDescConta = ""
                
        End If
         
        'Carrega ContaReceitaFinanceira na tela
        ContaReceitaFinanceira.PromptInclude = False
        ContaReceitaFinanceira.Text = sContaEnxutaRec
        ContaReceitaFinanceira.PromptInclude = True
        
        'Carrega a Descrição na tela
        DescContaReceita.Caption = objPlanoConta.sDescConta
    
    End If
    
    Traz_Dados_Tela_Contabilidade = SUCESSO
        
    Exit Function
    
Erro_Traz_Dados_Tela_Contabilidade:

    Traz_Dados_Tela_Contabilidade = Err
    
    Select Case Err
    
        Case 15149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objTiposDeAplicacao.sContaAplicacao)
        
        Case 15150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objTiposDeAplicacao.sContaReceita)
            
        Case 15343, 15345
        
        Case 15346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, objTiposDeAplicacao.sContaAplicacao)
            
        Case 15344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, objTiposDeAplicacao.sContaReceita)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174710)
            
    End Select
        
    Exit Function

End Function

Private Function Move_Tela_Memoria(objTiposDeAplicacao As ClassTiposDeAplicacao) As Long
'Move os dados da tela para objTiposDeAplicacao

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Se o Código do Tipo de aplicação não está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then
        
        'Inicializa o código no Obj
        objTiposDeAplicacao.iCodigo = 0
    
    'Se o Código do Tipo de aplicação está preenchido
    Else
        
        'Passa o Código do Tipo de aplicação que está na tela para o Obj
        objTiposDeAplicacao.iCodigo = CInt(Codigo.Text)
    
    End If

    'Passa os dados do Tipo de aplicação que estão na tela para o Obj
    objTiposDeAplicacao.sDescricao = Descricao.Text
    objTiposDeAplicacao.sHistorico = HistoricoPadrao.Text
    objTiposDeAplicacao.iInativo = Inativo.Value
    
    'Move os dados referentes à Contabilidade para a memória
    lErro = Move_Tela_Memoria_Contabil(objTiposDeAplicacao)
    If lErro <> SUCESSO Then Error 15168
        
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err
    
        Case 15168
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174711)

    End Select
    
    Exit Function

End Function

Private Function Move_Tela_Memoria_Contabil(objTiposDeAplicacao As ClassTiposDeAplicacao)
'Move os dados da tela referentes à Contabilidade para objTiposDeAplicacao

Dim lErro As Long
Dim iContaAplicPreenchida As Integer
Dim iContaRecPreenchida As Integer
Dim sContaAplicacao As String
Dim sContaReceita As String

On Error GoTo Erro_Move_Tela_Memoria_Contabil
 
    'Se o módulo de Contabilidade estiver ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        'Formata a Conta Aplicação que está na tela
        lErro = CF("Conta_Formata", ContaContabilAplicacao.Text, sContaAplicacao, iContaAplicPreenchida)
        If lErro <> SUCESSO Then Error 15121
    
        'Formata a Conta Receita que está na tela
        lErro = CF("Conta_Formata", ContaReceitaFinanceira.Text, sContaReceita, iContaRecPreenchida)
        If lErro <> SUCESSO Then Error 15122
        
        'Se a Conta Aplicação não está vazia
        If iContaAplicPreenchida <> CONTA_VAZIA Then
            
            'Passa a Conta Aplicação formatada para o Obj
            objTiposDeAplicacao.sContaAplicacao = sContaAplicacao
        
        'Se a Conta Aplicação está vazia
        Else
                    
            'Preenche a Conta Aplicação do Obj com vazio
            objTiposDeAplicacao.sContaAplicacao = ""
        
        End If
                
        'Se a Conta Receita não está vazia
        If iContaRecPreenchida <> CONTA_VAZIA Then
            
            'Passa a Conta Receita formatada para o Obj
            objTiposDeAplicacao.sContaReceita = sContaReceita
        
        'Se a Conta Receita está vazia
        Else
        
            'Preenche a Conta Receita do Obj com vazio
            objTiposDeAplicacao.sContaReceita = ""
            
        End If
            
    'Se o módulo de Contabilidade não estiver ativo
    Else
        
        'Preenche as Contas do Obj com vazio
        objTiposDeAplicacao.sContaAplicacao = ""
        objTiposDeAplicacao.sContaReceita = ""
        
    End If
    
    Move_Tela_Memoria_Contabil = SUCESSO

    Exit Function
    
Erro_Move_Tela_Memoria_Contabil:

    Move_Tela_Memoria_Contabil = Err
    
    Select Case Err
    
        Case 15121, 15122
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174712)
    
    End Select
    
    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TIPOS_APLICACAO
    Set Form_Load_Ocx = Me
    Caption = "Tipos de Aplicação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoAplicacao"
    
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

Private Sub Unload(objme As Object)
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is ContaContabilAplicacao Then
            Call LabelContaContabilAplicacao_Click
        ElseIf Me.ActiveControl Is ContaReceitaFinanceira Then
            Call LabelContaReceitaFinanceira_Click
        End If
    
    End If
    
End Sub

Function TipoAplicacao_Automatico(iCodigo As Integer) As Long
'Retorna o número da proximo Tipo de aplicação disponivel

Dim lCodigo As Long, lErro As Long

On Error GoTo Erro_TipoAplicacao_Automatico

    lErro = CF("Config_ObterAutomatico", "CPRConfig", NUM_PROX_TIPO_APLICACAO, "TiposDeAplicacao", "Codigo", lCodigo)
    If lErro <> SUCESSO Then Error 57751
    
    iCodigo = lCodigo
    
    TipoAplicacao_Automatico = SUCESSO

    Exit Function

Erro_TipoAplicacao_Automatico:

    TipoAplicacao_Automatico = Err

    Select Case Err

        Case 57751
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174713)

    End Select
    
    Exit Function
    
End Function


Private Sub LabelContaReceitaFinanceira_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaReceitaFinanceira, Source, X, Y)
End Sub

Private Sub LabelContaReceitaFinanceira_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaReceitaFinanceira, Button, Shift, X, Y)
End Sub

Private Sub LabelContaContabilAplicacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaContabilAplicacao, Source, X, Y)
End Sub

Private Sub LabelContaContabilAplicacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaContabilAplicacao, Button, Shift, X, Y)
End Sub

Private Sub DescContaContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescContaContabil, Source, X, Y)
End Sub

Private Sub DescContaContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescContaContabil, Button, Shift, X, Y)
End Sub

Private Sub LabelDescContaContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescContaContabil, Source, X, Y)
End Sub

Private Sub LabelDescContaContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescContaContabil, Button, Shift, X, Y)
End Sub

Private Sub DescContaReceita_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescContaReceita, Source, X, Y)
End Sub

Private Sub DescContaReceita_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescContaReceita, Button, Shift, X, Y)
End Sub

Private Sub LabelDescContaReceita_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescContaReceita, Source, X, Y)
End Sub

Private Sub LabelDescContaReceita_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescContaReceita, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelTiposDeAplicacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTiposDeAplicacao, Source, X, Y)
End Sub

Private Sub LabelTiposDeAplicacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTiposDeAplicacao, Button, Shift, X, Y)
End Sub

Private Sub LabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistoricos, Source, X, Y)
End Sub

Private Sub LabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub LabelPlanoDeContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPlanoDeContas, Source, X, Y)
End Sub

Private Sub LabelPlanoDeContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPlanoDeContas, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

