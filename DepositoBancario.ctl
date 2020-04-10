VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl DepositoBancario 
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   KeyPreview      =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   7530
   Begin VB.Frame Frame2 
      Caption         =   "Dados Principais"
      Height          =   1320
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   720
      Width           =   7275
      Begin VB.ComboBox CodContaCorrente 
         Height          =   315
         Left            =   1665
         TabIndex        =   1
         Top             =   330
         Width           =   1695
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   6315
         Picture         =   "DepositoBancario.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Numeração Automática"
         Top             =   330
         Width           =   300
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   1665
         TabIndex        =   4
         Top             =   810
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Sequencial 
         Height          =   300
         Left            =   5220
         TabIndex        =   2
         Top             =   315
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   2820
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   810
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   5205
         TabIndex        =   6
         Top             =   810
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   4605
         TabIndex        =   20
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   1095
         TabIndex        =   19
         Top             =   840
         Width           =   480
      End
      Begin VB.Label LabelSequencial 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial:"
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
         Left            =   4095
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label LblConta 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente:"
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   375
         Width           =   1350
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Complemento"
      Height          =   795
      Left            =   90
      TabIndex        =   21
      Top             =   2130
      Width           =   7275
      Begin MSMask.MaskEdBox Historico 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   285
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Histórico:"
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
         Left            =   795
         TabIndex        =   22
         Top             =   345
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5220
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "DepositoBancario.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "DepositoBancario.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "DepositoBancario.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "DepositoBancario.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa Central"
      Height          =   870
      Left            =   90
      TabIndex        =   0
      Top             =   3045
      Width           =   7275
      Begin VB.OptionButton CaixaCentral 
         Caption         =   "Central"
         Enabled         =   0   'False
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
         Left            =   -20000
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton CaixaGeral 
         Caption         =   "Geral"
         Enabled         =   0   'False
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
         Left            =   -20000
         TabIndex        =   13
         Top             =   570
         Width           =   1050
      End
      Begin VB.Label SaldoDinheiro 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3795
         TabIndex        =   15
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo em dinheiro:"
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
         Left            =   2055
         TabIndex        =   14
         Top             =   375
         Width           =   1590
      End
   End
End
Attribute VB_Name = "DepositoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Buttons CaixaCentral e CaixaGeral são meramente informativos. Não
'podem ser editados.

Private WithEvents objEventoDeposito As AdmEvento
Attribute objEventoDeposito.VB_VarHelpID = -1
Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        Call LabelSequencial_Click
    End If

End Sub

Public Sub form_unload(Cancel As Integer)

On Error GoTo Erro_Form_Unload
    
    'libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)
    Set objEventoDeposito = Nothing
    Set objEventoContaCorrenteInt = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158827)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub Limpa_Tela_Deposito()

Dim lErro As Long
Dim dValor As Double
Dim objTipoMeioPagtoLoja As New ClassTMPLojaFilial

On Error GoTo Erro_Limpa_Tela_Deposito

    'limpa os campos da tela
    Call Limpa_Tela(Me)
    
    CodContaCorrente.ListIndex = -1
    
    'preenche o tipo meio pagamento comm pagamento dinheiro
    objTipoMeioPagtoLoja.iTipo = MEIO_PAGAMENTO_DINHEIRO
    objTipoMeioPagtoLoja.iFilialEmpresa = giFilialEmpresa
    
    'le o saldo atualizado da tabela
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTipoMeioPagtoLoja)
    If lErro <> SUCESSO Then gError 103896
    
    'preenche o campo valor com o saldo
    SaldoDinheiro.Caption = Format(objTipoMeioPagtoLoja.dSaldo, "STANDARD")
    
    'preenche o campo data com a data atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    Exit Sub

Erro_Limpa_Tela_Deposito:

    Select Case gErr
    
        Case 103896
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158828)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'testa se houve alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 103932

    'limpa os campos que se fizerem necessários
    Call Limpa_Tela_Deposito

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 103933

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 103932, 103933
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158829)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objLog As New ClassLog
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim objCCMovDia As New ClassCCMovDia

On Error GoTo Erro_BotaoExcluir_Click

    'se estiver no BO-> erro
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 126041

    'verifica se sequencial está preenchido
    If Len(Trim(Sequencial.Text)) = 0 Then gError 103926

    objMovimentoCaixa.iCaixa = CODIGO_CAIXA_CENTRAL
    objMovimentoCaixa.lSequencial = StrParaLong(Sequencial.Text)

    lErro = CF("MovimentosCaixa_Le", objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103744 Then gError 103927

    If lErro = 103744 Then gError 103928

    'se o movimento não for do tipo Deposito--> erro
    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_DEPOSITO_BANCARIO Then gError 103929
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTOSCAIXA_EXCLUSAO", STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)
    
    If vbMsgRes = vbYes Then

        objLog.iOperacao = EXCLUSAO_DEPOSITO_BANCARIO

        lErro = Mover_Dados_DepositoBancario_CCMovDia(objMovimentoCaixa, objCCMovDia)
        If lErro <> SUCESSO Then gError 103930

        Call Mover_Dados_DepositoBancario_Log_Exc(objMovimentoCaixa, objLog)
        
        lErro = CF("MovimentosCaixa_Exclui", objMovimentoCaixa, objLog, objCCMovDia)
        If lErro <> SUCESSO Then gError 103931

        Call Limpa_Tela_Deposito

        iAlterado = 0

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 103926
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)

        Case 103927, 103930, 103931

        Case 103928
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOSCAIXA_NAOENCONTRADO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case 103929
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO_BANCARIO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case 126041
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NAO_PERMITIDA_BACKOFFICE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158830)

    End Select

    Exit Sub

End Sub

Private Sub Mover_Dados_DepositoBancario_Log_Exc(objMovimentoCaixa As ClassMovimentoCaixa, objLog As ClassLog)
'Função que recebe um objMovimentoCaixa com todos os dados preenchidos e preenche a strin de um
'objLog para prepará-lo para a gravação

On Error GoTo Erro_Mover_Dados_DepositoBancario_Log_Exc

    'preenche os dados do log com cada atributo do objMovimentoCaixa separado por vbkeyscape
    With objMovimentoCaixa
        objLog.sLog = CStr(.iFilialEmpresa) & Chr(vbKeyEscape) & _
                      CStr(.iCaixa) & Chr(vbKeyEscape) & _
                      CStr(.lSequencial) & Chr(vbKeyEscape) & _
                      Chr(vbKeyEnd)
    End With

    Exit Sub

Erro_Mover_Dados_DepositoBancario_Log_Exc:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158831)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que preenche os dados de um objMovimentoCaixa com os dados da tela

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'preenche os atributos do obj
    objMovimentoCaixa.iFilialEmpresa = giFilialEmpresa
    objMovimentoCaixa.iCaixa = CODIGO_CAIXA_CENTRAL
    objMovimentoCaixa.lSequencial = StrParaLong(Sequencial.Text)
    objMovimentoCaixa.iTipo = MOVIMENTO_CAIXA_DEPOSITO_BANCARIO
    objMovimentoCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
    objMovimentoCaixa.iParcelamento = PARCELAMENTO_AVISTA
    objMovimentoCaixa.dtDataMovimento = StrParaDate(Data.Text)
    objMovimentoCaixa.dValor = StrParaDbl(Valor.Text)
    objMovimentoCaixa.sHistorico = Trim(Historico.Text)
    objMovimentoCaixa.iCodConta = Codigo_Extrai(CodContaCorrente.List(CodContaCorrente.ListIndex))
    objMovimentoCaixa.dHora = CDbl(Time)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158832)

    End Select

    Exit Function

End Function

Private Function Traz_Deposito_Tela(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que preenche a tela com os dados selecionados no browser

Dim lErro As Long
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Traz_Deposito_Tela
    
    Call Limpa_Tela_Deposito
    
    'busca o movimento de caixa
    lErro = CF("MovimentosCaixa_Le_NumIntDoc", objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103677 Then gError 103871

    If lErro = 103677 Then gError 103872

    'verifica se o tipo de movimento é depósito Bancário
    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_DEPOSITO_BANCARIO Then gError 103873

    'preenche a tela
    Sequencial.Text = objMovimentoCaixa.lSequencial
    Data.PromptInclude = False
    Data.Text = Format(objMovimentoCaixa.dtDataMovimento, "dd/mm/yy")
    Data.PromptInclude = True
    Valor.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
    Historico.Text = objMovimentoCaixa.sHistorico
    CodContaCorrente.Text = objMovimentoCaixa.iCodConta
    Call CodContaCorrente_Validate(bCancel)
    
'    For iIndice = 0 To CodContaCorrente.ListCount - 1
'
'        If CodContaCorrente.ItemData(iIndice) = objMovimentoCaixa.iCodConta Then
'
'            CodContaCorrente.ListIndex = iIndice
'            Exit For
'
'        End If
'
'    Next

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 103874

    Traz_Deposito_Tela = SUCESSO

    Exit Function

Erro_Traz_Deposito_Tela:

    Traz_Deposito_Tela = gErr

        Select Case gErr

        Case 103875

        Case 103871, 103872, 103874

        Case 103873
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO_BANCARIO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158833)

    End Select

    Exit Function

End Function

Private Function Carrega_Conta_Corrente() As Long

Dim colContasCorrentesInternas As New Collection
Dim lErro As Long
Dim objContaCorrenteInterna As ClassContasCorrentesInternas

On Error GoTo Erro_Carrega_Conta_Corrente

    'carrega a coleção com os dados da tabela
    lErro = CF("ContasCorrentesInternas_Le_Todas", colContasCorrentesInternas)
    If lErro <> SUCESSO And lErro <> 103865 Then gError 103867
    
    'se a tabela estiver vazia -> erro
    If lErro = 103865 Then gError 103868
    
    For Each objContaCorrenteInterna In colContasCorrentesInternas
    
        'preenche um item da combo e o itemdata do mesmo
        CodContaCorrente.AddItem (objContaCorrenteInterna.iCodigo & SEPARADOR & objContaCorrenteInterna.sNomeReduzido)
        CodContaCorrente.ItemData(CodContaCorrente.NewIndex) = objContaCorrenteInterna.iCodigo
    
    Next

    Carrega_Conta_Corrente = SUCESSO
    
    Exit Function

Erro_Carrega_Conta_Corrente:
    
    Carrega_Conta_Corrente = gErr
    
    Select Case gErr
    
        Case 103867
        
        Case 103868
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_VAZIA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158834)
    
    End Select
    
    Exit Function

End Function

Private Function Mover_Dados_DepositoBancario_CCMovDia(objMovimentoCaixa As ClassMovimentoCaixa, objCCMovDia As ClassCCMovDia) As Long
'Função que recebe um objMovimentoCaixa com todos os dados preenchidos e preenche o objCCmovDia para sua atualalização
    
On Error GoTo Erro_Mover_Dados_DepositoBancario_CCMovDia

    'preenche os dados do objCCMovDia com os dados do movimento de caixa passado por parâmetro
    objCCMovDia.dDeb = objMovimentoCaixa.dValor
    objCCMovDia.dtData = objMovimentoCaixa.dtDataMovimento
    objCCMovDia.iCodCaixa = objMovimentoCaixa.iCaixa
    objCCMovDia.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto
    objCCMovDia.iFilialEmpresa = objMovimentoCaixa.iFilialEmpresa
    objCCMovDia.iParcelamento = objMovimentoCaixa.iParcelamento
    objCCMovDia.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_DINHEIRO

    Mover_Dados_DepositoBancario_CCMovDia = SUCESSO

    Exit Function

Erro_Mover_Dados_DepositoBancario_CCMovDia:

    Mover_Dados_DepositoBancario_CCMovDia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158835)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro()

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se estiver no BO-> erro
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 126040
    
    'verifica se a conta foi selecionada
    If CodContaCorrente.ListIndex = -1 Then gError 103895
    
    'verifica se o seqüencial está preenchido
    If Len(Trim(Sequencial.Text)) = 0 Then gError 103890

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 103891

    'carrega o objCaixa com os campos chave
    objCaixa.iFilialEmpresa = giFilialEmpresa
    objCaixa.iCodigo = CODIGO_CAIXA_CENTRAL

    'verifica se o valor está preenchido
    If Len(Trim(Valor.Text)) = 0 Then gError 103892

    'preenche o obj com os dados da tela
    lErro = Move_Tela_Memoria(objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103893
    
    'Pergunta se deseja alterar
    lErro = Trata_Alteracao(objMovimentoCaixa, objMovimentoCaixa.iCaixa, objMovimentoCaixa.lSequencial)
    If lErro <> SUCESSO Then gError 107047
    
    'grava o Deposito Bancario
    lErro = CF("Movimentos_Caixa_Grava_DepositoBancario", objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103894

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 103890
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", gErr)

        Case 103891
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 103892
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DEPOSITO_NAO_PREENCHIDO", gErr)

        Case 103893, 103894, 107047
        
        Case 103895
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_SELECIONADA", gErr)
        
        Case 126040
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_NAO_PERMITIDA_BACKOFFICE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158836)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'grava o registro
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 103889

    'limpa a tela
    Call Limpa_Tela_Deposito

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 103889

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158837)

    End Select

    Exit Sub

End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objContaCorrenteInterna As New ClassContasCorrentesInternas
Dim bCancel As Boolean

On Error GoTo Erro_objEventoContaCorrenteInt_evSelecao

    Set objContaCorrenteInterna = obj1
    
    CodContaCorrente.Text = objContaCorrenteInterna.iCodigo
    
    Call CodContaCorrente_Validate(bCancel)
    
    Me.Show
    
    Exit Sub

Erro_objEventoContaCorrenteInt_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158838)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'se o campo valor estiver preenchido
    If Len(Trim(Valor.Text)) <> 0 Then

        'critica o valor
        lErro = Valor_Positivo_Critica(Trim(Valor.Text))
        If lErro <> SUCESSO Then gError 103888

    End If
    
    Cancel = False

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 103888

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158839)

    End Select

    Exit Sub

End Sub

Private Sub Valor_GotFocus()

On Error GoTo Erro_Valor_GotFocus

    Call MaskEdBox_TrataGotFocus(Valor, iAlterado)

    Exit Sub

Erro_Valor_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158840)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'aumenta a data de um dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 103887

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 103887

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158841)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'diminui a data de um dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 103886

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 103886

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158842)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'se o campo data não etiver preenchido, sai
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    'critica o dado do campo
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 103885
    
    Cancel = False

    Exit Sub

Erro_Data_Validate:

    Cancel = True
    
    Select Case gErr

        Case 103885

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158843)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()

On Error GoTo Erro_Data_GotFocus

    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

    Exit Sub

Erro_Data_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158844)

    End Select

Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lSequencial As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_BotaoProxNum_Click

    'preenche o objMovimentoCaixa com os dado necessários para uma chave candidata
    objMovimentoCaixa.iFilialEmpresa = giFilialEmpresa
    objMovimentoCaixa.iCaixa = CODIGO_CAIXA_CENTRAL

    'chama a função que gerará o próximo número
    lErro = CF("Caixa_Sequencial_Transacao", objMovimentoCaixa.iCaixa, objMovimentoCaixa.iFilialEmpresa, lSequencial)
    If lErro <> SUCESSO Then gError 103884

    Sequencial.Text = lSequencial

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 103884

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158845)

    End Select

    Exit Sub

End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)

On Error GoTo Erro_Sequencial_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(Sequencial.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(Sequencial.Text) Then gError 103882

        'Verifica se codigo é menor que um
        If StrParaLong(Sequencial.Text) < 1 Then gError 103883

    End If

    Cancel = False

    Exit Sub

Erro_Sequencial_Validate:

    Cancel = True

    Select Case gErr

        Case 103882, 103883
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_INVALIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158846)

    End Select

    Exit Sub

End Sub

Private Sub Sequencial_GotFocus()

Dim lErro As Long

On Error GoTo Erro_Sequencial_GotFocus

    Call MaskEdBox_TrataGotFocus(Sequencial, iAlterado)

Exit Sub

Erro_Sequencial_GotFocus:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158847)

    End Select

Exit Sub

End Sub

Private Sub Sequencial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objContaCorrenteInterna As New ClassContasCorrentesInternas

On Error GoTo Erro_CodContaCorrente_Validate

    'verifica se a combo esta preenchida
    If Len(Trim(CodContaCorrente.Text)) = 0 Then Exit Sub
    
    'verifica se a combo foi selecionada
    If CodContaCorrente.ListIndex <> -1 Then Exit Sub
    
    'tenta selecionar o item da combo
    lErro = Combo_Seleciona(CodContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 103934
    
    'se nao achou pelo codigo
    If lErro = 6730 Then
    
        'preenche a chave de busca da conta corrente
        objContaCorrenteInterna.iCodigo = iCodigo
        
        'tenta buscar no banco de dados
        lErro = CF("ContasCorrentesInternas_Le", objContaCorrenteInterna)
        If lErro <> SUCESSO And lErro <> 103938 Then gError 103939
        
        'se não encontrou -> erro
        If lErro = 103938 Then gError 103940
             
        'se encontrou, preenche a combo com o indivíduo
        CodContaCorrente.Text = CStr(objContaCorrenteInterna.iCodigo) & SEPARADOR & objContaCorrenteInterna.sNomeReduzido
        
    End If
    
    'se não achou pela string -> erro
    If lErro = 6731 Then gError 103941
    
    Cancel = False
    
    Exit Sub

Erro_CodContaCorrente_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103934, 103939
        
        Case 103940
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_NAOENCONTRADA", gErr, objContaCorrenteInterna.iCodigo)
            
        Case 103941
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTASCORRENTESINTERNAS_NAOENCONTRADA", gErr, CodContaCorrente.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158848)
    
    End Select
    
    Exit Sub

End Sub


Private Sub CodContaCorrente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Tela_Preenche

    'preenche o objmovimentocaixa com a colecao de valores
    objMovimentoCaixa.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objMovimentoCaixa.lNumMovto = colCampoValor.Item("NumMovto").vValor

    'traz os dados do Deposito Bancario para a tela
    lErro = Traz_Deposito_Tela(objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103654
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 103654
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158849)

    End Select

    Exit Sub

End Sub

Public Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Tela_Extrai

    sTabela = "MovimentosCaixa"

    'preenche o objMovimentoCaixa com os dados da tela
    lErro = Move_Tela_Memoria(objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103880

    'preenche a coleção de campos-valores
    colCampoValor.Add "FilialEmpresa", objMovimentoCaixa.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumMovto", objMovimentoCaixa.lNumMovto, 0, "NumMovto"
    colCampoValor.Add "Caixa", objMovimentoCaixa.iCaixa, 0, "Caixa"
    colCampoValor.Add "Sequencial", objMovimentoCaixa.lSequencial, 0, "Sequencial"
    colCampoValor.Add "Tipo", objMovimentoCaixa.iTipo, 0, "Tipo"
    colCampoValor.Add "AdmMeioPagto", objMovimentoCaixa.iAdmMeioPagto, 0, "AdmMeioPagto"
    colCampoValor.Add "Parcelamento", objMovimentoCaixa.iParcelamento, 0, "Parcelamento"
    colCampoValor.Add "TipoCartao", objMovimentoCaixa.iTipoCartao, 0, "TipoCartao"
    colCampoValor.Add "Numero", objMovimentoCaixa.lNumero, 0, "Numero"
    colCampoValor.Add "DataMovimento", objMovimentoCaixa.dtDataMovimento, 0, "DataMovimento"
    colCampoValor.Add "Valor", objMovimentoCaixa.dValor, 0, "Valor"
    colCampoValor.Add "Historico", objMovimentoCaixa.sHistorico, STRING_MOVIMENTOCAIXA_HISTORICO, "Historico"
    colCampoValor.Add "Favorecido", objMovimentoCaixa.sFavorecido, STRING_MOVIMENTOCAIXA_FAVORECIDO, "Favorecido"
    colCampoValor.Add "CupomFiscal", objMovimentoCaixa.lCupomFiscal, 0, "CupomFiscal"
    colCampoValor.Add "NumRefInterna", objMovimentoCaixa.lNumRefInterna, 0, "NumRefInterna"
    colCampoValor.Add "MovtoTransf", objMovimentoCaixa.lMovtoTransf, 0, "MovtoTransf"
    colCampoValor.Add "MovtoEstorno", objMovimentoCaixa.lMovtoEstorno, 0, "MovtoEstorno"
    colCampoValor.Add "Gerente", objMovimentoCaixa.iGerente, 0, "Gerente"
    colCampoValor.Add "CodConta", objMovimentoCaixa.iCodConta, 0, "CodConta"

    'estabelece o filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Tipo", OP_IGUAL, MOVIMENTO_CAIXA_DEPOSITO_BANCARIO

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 103880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158850)

    End Select

    Exit Function

End Function

Private Sub objEventoDeposito_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovimentoCaixa As ClassMovimentoCaixa

On Error GoTo Erro_objEventoDeposito_evSelecao

    Set objMovimentoCaixa = obj1
    
    'traz os dados do Deposito Bancario para a tela
    lErro = Traz_Deposito_Tela(objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103872 Then gError 103654
    
    If lErro = 103872 Then gError 103878
    
'    'tenta ler do bd
'    lErro = CF("MovimentosCaixa_Le", objMovimentoCaixa)
'    If lErro <> SUCESSO And lErro <> 103744 Then gError 103877
'
'    'se não encontrou-> erro
'    If lErro = 103744 Then gError 103878
'
'    'se o movimento encontrado no bd não for do tipo Depósito Bancário, erro
'    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_DEPOSITO_BANCARIO Then gError 103879
'
'    'mover os dados para a tela
'    Sequencial.Text = objMovimentoCaixa.lSequencial
'
'    Data.PromptInclude = False
'    Data.Text = Format(objMovimentoCaixa.dtDataMovimento, "dd/mm/yy")
'    Data.PromptInclude = True
'
'    Valor.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
'    Historico.Text = objMovimentoCaixa.sHistorico
    
    iAlterado = 0
    
    'mostrar a tela
    Me.Show

    Exit Sub

Erro_objEventoDeposito_evSelecao:

    Select Case gErr
    
        Case 103877
        
        Case 103878
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOSCAIXA_NAOENCONTRADO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)
            
        Case 103879
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO_BANCARIO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158851)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelSequencial_Click()

Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim colSelecao As New Collection

On Error GoTo Erro_LabelSequencial_Click

    'se o sequencial estiver preenchido
    If Len(Trim(Sequencial.Text)) <> 0 Then

        'preenche o atributo seqüencial do obj com o conteúdo do campo seqüencial
        objMovimentoCaixa.lSequencial = StrParaLong(Trim(Sequencial.Text))

    End If

    Call Chama_Tela("DepositoBancarioLista", colSelecao, objMovimentoCaixa, objEventoDeposito)

    Exit Sub

Erro_LabelSequencial_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158852)

    End Select

    Exit Sub

End Sub

Private Sub LblConta_Click()

Dim objContaCorrenteInterna As New ClassContasCorrentesInternas
Dim colSelecao As New Collection

On Error GoTo Erro_lblConta_Click

    'se a conta estiver preenchida
    If CodContaCorrente.ListIndex <> -1 Then

        'preenche o codigo do obj com o conteúdo da combo
        objContaCorrenteInterna.iCodigo = Codigo_Extrai(CodContaCorrente.List(CodContaCorrente.ListIndex))

    End If

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContaCorrenteInterna, objEventoContaCorrenteInt)

    Exit Sub

Erro_lblConta_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158853)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objMovimentoCaixa As ClassMovimentoCaixa) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se há movimento de caixa
    If Not (objMovimentoCaixa Is Nothing) Then

        'traz os dados do depósito bancário para a tela
        lErro = Traz_Deposito_Tela(objMovimentoCaixa)
        If lErro <> SUCESSO And lErro <> 103872 Then gError 103876

        'se retornou erro indicando que não está cadastrado
        If lErro = 103872 Then

            'limpa a tela
            Call Limpa_Tela_Deposito

            'coloca o seqüencial na tela
            Sequencial.Text = objMovimentoCaixa.lSequencial

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 103876

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158854)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim dValor As Double
Dim objTipoMeioPagtoLoja As New ClassTMPLojaFilial

On Error GoTo Erro_Form_Load

    'seta o admEvento
    Set objEventoDeposito = New AdmEvento
    Set objEventoContaCorrenteInt = New AdmEvento

    'preenche o campo data com a data atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    objTipoMeioPagtoLoja.iTipo = MEIO_PAGAMENTO_DINHEIRO
    objTipoMeioPagtoLoja.iFilialEmpresa = giFilialEmpresa
    
    'consulda o saldo
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTipoMeioPagtoLoja)
    If lErro <> SUCESSO Then gError 103869

    lErro = Carrega_Conta_Corrente
    If lErro <> SUCESSO Then gError 103871

    'preenche o campo valor com o saldo
    SaldoDinheiro.Caption = Format(objTipoMeioPagtoLoja.dSaldo, "STANDARD")

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103869, 103871

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158855)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    ' Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Depósito Bancário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "DepositoBancario"
    
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

Private Sub GerenteSenha_Change()

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
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
