VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl DepositoCaixa 
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   7470
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5205
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "DepositoCaixa.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "DepositoCaixa.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "DepositoCaixa.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "DepositoCaixa.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa Central"
      Height          =   825
      Left            =   60
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
         TabIndex        =   19
         TabStop         =   0   'False
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   570
         Width           =   1050
      End
      Begin VB.Label SaldoDinheiro 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2970
         TabIndex        =   13
         Top             =   315
         Width           =   1890
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
         Left            =   1230
         TabIndex        =   12
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Principais"
      Height          =   1320
      Left            =   90
      TabIndex        =   14
      Top             =   675
      Width           =   7275
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2895
         Picture         =   "DepositoCaixa.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   315
         Width           =   300
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   4800
         TabIndex        =   4
         Top             =   315
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   5955
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   1800
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
      Begin MSMask.MaskEdBox Sequencial 
         Height          =   300
         Left            =   1800
         TabIndex        =   1
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
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
         Left            =   675
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   345
         Width           =   1020
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
         Left            =   4260
         TabIndex        =   3
         Top             =   375
         Width           =   480
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
         Left            =   1185
         TabIndex        =   15
         Top             =   870
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Complemento"
      Height          =   885
      Left            =   60
      TabIndex        =   16
      Top             =   2055
      Width           =   7275
      Begin MSMask.MaskEdBox Historico 
         Height          =   315
         Left            =   1455
         TabIndex        =   7
         Top             =   330
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
         Left            =   555
         TabIndex        =   17
         Top             =   390
         Width           =   825
      End
   End
End
Attribute VB_Name = "DepositoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis Globais
Dim iAlterado As Integer
Private WithEvents objEventoDeposito As AdmEvento
Attribute objEventoDeposito.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub form_unload(Cancel As Integer)

On Error GoTo Erro_Form_Unload
    
    'libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)
    Set objEventoDeposito = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158856)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

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
    
    'preenche o tipo meio pagamento comm pagamento dinheiro
    objTipoMeioPagtoLoja.iTipo = MEIO_PAGAMENTO_DINHEIRO
    objTipoMeioPagtoLoja.iFilialEmpresa = giFilialEmpresa
    
    'le o saldo atualizado da tabela
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTipoMeioPagtoLoja)
    If lErro <> SUCESSO Then gError 103855
    
    'preenche o campo valor com o saldo
    SaldoDinheiro.Caption = Format(objTipoMeioPagtoLoja.dSaldo, "STANDARD")
    
    'preenche o campo data com a data atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    Exit Sub

Erro_Limpa_Tela_Deposito:

    Select Case gErr
    
        Case 103855
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158857)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'testa se houve alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 103853

    'limpa os campos que se fizerem necessários
    Call Limpa_Tela_Deposito

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 103854

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 103853, 103854
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158858)

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
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 126043

    'verifica se sequencial está preenchido
    If Len(Trim(Sequencial.Text)) = 0 Then gError 103847

    objMovimentoCaixa.iCaixa = CODIGO_CAIXA_CENTRAL
    objMovimentoCaixa.lSequencial = StrParaLong(Sequencial.Text)

    lErro = CF("MovimentosCaixa_Le", objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103744 Then gError 103848

    If lErro = 103744 Then gError 103849

    'se o movimento não for do tipo Deposito--> erro
    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_DEPOSITO_DINHEIRO Then gError 103850
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTOSCAIXA_EXCLUSAO", STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)
    If vbMsgRes = vbYes Then

        objLog.iOperacao = EXCLUSAO_DEPOSITO_CAIXA_CENTRAL

        Call Mover_Dados_Deposito_Log_Exc(objMovimentoCaixa, objLog)

        lErro = Mover_Dados_Deposito_CCMovDia(objMovimentoCaixa, objCCMovDia)
        If lErro <> SUCESSO Then gError 103851

        'muda o sinal do deposito para a exclusão ocorrer corretamente
        objMovimentoCaixa.dValor = -objMovimentoCaixa.dValor

        lErro = CF("MovimentosCaixa_Exclui", objMovimentoCaixa, objLog, objCCMovDia)
        If lErro <> SUCESSO Then gError 103852

        Call Limpa_Tela_Deposito

        iAlterado = 0

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 103847
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)

        Case 103848, 103851, 103852

        Case 103849
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOSCAIXA_NAOENCONTRADO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case 103850
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO_DINHEIRO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case 126043
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NAO_PERMITIDA_BACKOFFICE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158859)

    End Select

    Exit Sub

End Sub

Private Sub Mover_Dados_Deposito_Log_Exc(objMovimentoCaixa As ClassMovimentoCaixa, objLog As ClassLog)
'Função que recebe um objMovimentoCaixa com todos os dados preenchidos e preenche a strin de um
'objLog para prepará-lo para a gravação

On Error GoTo Erro_Mover_Dados_Deposito_Log_Exc

    'preenche os dados do log com cada atributo do objMovimentoCaixa separado por vbkeyscape
    With objMovimentoCaixa
        objLog.sLog = CStr(.iFilialEmpresa) & Chr(vbKeyEscape) & _
                      CStr(.iCaixa) & Chr(vbKeyEscape) & _
                      CStr(.lSequencial) & Chr(vbKeyEscape) & _
                      Chr(vbKeyEnd)

    End With

    Exit Sub

Erro_Mover_Dados_Deposito_Log_Exc:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158860)

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
    objMovimentoCaixa.iTipo = MOVIMENTO_CAIXA_DEPOSITO_DINHEIRO
    objMovimentoCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
    objMovimentoCaixa.iParcelamento = PARCELAMENTO_AVISTA
    objMovimentoCaixa.dtDataMovimento = StrParaDate(Data.Text)
    objMovimentoCaixa.dValor = StrParaDbl(Valor.Text)
    objMovimentoCaixa.sHistorico = Trim(Historico.Text)
    objMovimentoCaixa.dHora = CDbl(Time)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158861)

    End Select

    Exit Function

End Function

Private Function Mover_Dados_Deposito_CCMovDia(objMovimentoCaixa As ClassMovimentoCaixa, objCCMovDia As ClassCCMovDia) As Long
'Função que recebe um objMovimentoCaixa com todos os dados preenchidos e preenche o objCCmovDia para sua atualalização
    
On Error GoTo Erro_Mover_Dados_Deposito_CCMovDia

    'preenche os dados do objCCMovDia com os dados do movimento de caixa passado por parâmetro
    objCCMovDia.dCred = objMovimentoCaixa.dValor
    objCCMovDia.dtData = objMovimentoCaixa.dtDataMovimento
    objCCMovDia.iCodCaixa = objMovimentoCaixa.iCaixa
    objCCMovDia.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto
    objCCMovDia.iFilialEmpresa = objMovimentoCaixa.iFilialEmpresa
    objCCMovDia.iParcelamento = objMovimentoCaixa.iParcelamento
    objCCMovDia.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_DINHEIRO

    Mover_Dados_Deposito_CCMovDia = SUCESSO

    Exit Function

Erro_Mover_Dados_Deposito_CCMovDia:

    Mover_Dados_Deposito_CCMovDia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158862)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se estiver no BO-> erro
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 126042

    'verifica se o seqüencial está preenchido
    If Len(Trim(Sequencial.Text)) = 0 Then gError 103822

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 103823

    'carrega o objCaixa com os campos chave
    objCaixa.iFilialEmpresa = giFilialEmpresa
    objCaixa.iCodigo = CODIGO_CAIXA_CENTRAL

    'verifica se o valor está preenchido
    If Len(Trim(Valor.Text)) = 0 Then gError 103824

    'preenche o obj com os dados da tela
    lErro = Move_Tela_Memoria(objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103825

    'Pergunta se deseja alterar
    lErro = Trata_Alteracao(objMovimentoCaixa, objMovimentoCaixa.iCaixa, objMovimentoCaixa.lSequencial)
    If lErro <> SUCESSO Then gError 107048

    'grava o Deposito
    lErro = CF("Movimentos_Caixa_Grava_DepositoCaixa", objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103826

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 103822
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", gErr)

        Case 103823
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 103824
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DEPOSITO_NAO_PREENCHIDO", gErr)

        Case 103825, 103826, 107048

        Case 126042
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_NAO_PERMITIDA_BACKOFFICE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158863)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'grava o registro
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 103821

    'limpa a tela
    Call Limpa_Tela_Deposito

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 103821

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158864)

    End Select

    Exit Sub

End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        Call LabelSequencial_Click
    End If

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
        If lErro <> SUCESSO Then gError 103820

    End If
    
    Cancel = False

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 103820

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158865)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158866)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'aumenta a data de um dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 103819

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 103819

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158867)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'diminui a data de um dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 103818

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 103818

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158868)

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
    If lErro <> SUCESSO Then gError 103817
    
    Cancel = False

    Exit Sub

Erro_Data_Validate:

    Cancel = True
    
    Select Case gErr

        Case 103817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158869)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158870)

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
    If lErro <> SUCESSO Then gError 103816

    Sequencial.Text = lSequencial

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 103816

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158871)

    End Select

    Exit Sub

End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)

On Error GoTo Erro_Sequencial_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(Sequencial.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(Sequencial.Text) Then gError 103814

        'Verifica se codigo é menor que um
        If StrParaLong(Sequencial.Text) < 1 Then gError 103815

    End If

    Cancel = False

    Exit Sub

Erro_Sequencial_Validate:

    Cancel = True

    Select Case gErr

        Case 103814, 103815
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_INVALIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158872)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158873)

    End Select

    Exit Sub

End Sub

Private Sub Sequencial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Tela_Preenche

    'preenche o objmovimentocaixa com a colecao de valores
    objMovimentoCaixa.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objMovimentoCaixa.lNumMovto = colCampoValor.Item("NumMovto").vValor

    'traz os dados do Depósito para a tela
    lErro = Traz_Deposito_Tela(objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103812
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 103812
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158874)

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
    If lErro <> SUCESSO Then gError 103811

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
    colSelecao.Add "Tipo", OP_IGUAL, MOVIMENTO_CAIXA_DEPOSITO_DINHEIRO

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 103811

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158875)

    End Select

    Exit Function

End Function

Private Sub objEventoDeposito_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovimentoCaixa As ClassMovimentoCaixa

On Error GoTo Erro_objEventoDeposito_evSelecao

    Set objMovimentoCaixa = obj1
    
    'tenta ler do bd
    lErro = CF("MovimentosCaixa_Le", objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103744 Then gError 103808
    
    'se não encontrou-> erro
    If lErro = 103744 Then gError 103809
    
    'se o movimento encontrado no bd não for do tipo Depósito, erro
    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_DEPOSITO_DINHEIRO Then gError 103810
    
    'mover os dados para a tela
    Sequencial.Text = objMovimentoCaixa.lSequencial
    
    Data.PromptInclude = False
    Data.Text = Format(objMovimentoCaixa.dtDataMovimento, "dd/mm/yy")
    Data.PromptInclude = True
    Valor.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
    Historico.Text = objMovimentoCaixa.sHistorico
    
    iAlterado = 0
    
    'mostrar a tela
    Me.Show

    Exit Sub

Erro_objEventoDeposito_evSelecao:

    Select Case gErr
    
        Case 103808
        
        Case 103809
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOSCAIXA_NAOENCONTRADO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)
            
        Case 103810
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO_DINHEIRO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158876)

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

    Call Chama_Tela("DepositoLojaLista", colSelecao, objMovimentoCaixa, objEventoDeposito)

    Exit Sub

Erro_LabelSequencial_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158877)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objMovimentoCaixa As ClassMovimentoCaixa) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se há movimento de caixa
    If Not (objMovimentoCaixa Is Nothing) Then

        'traz os dados do Depósito para a tela
        lErro = Traz_Deposito_Tela(objMovimentoCaixa)
        If lErro <> SUCESSO And lErro <> 103803 Then gError 103807

        'se retornou erro indicando que não está cadastrado
        If lErro = 103803 Then

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

        Case 103807

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158878)

    End Select

    Exit Function

End Function

Private Function Traz_Deposito_Tela(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que preenche a tela com os dados selecionados no browser

Dim lErro As Long

On Error GoTo Erro_Traz_Deposito_Tela

    'busca o movimento de caixa
    lErro = CF("MovimentosCaixa_Le_NumIntDoc", objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103677 Then gError 103802

    If lErro = 103677 Then gError 103803

    'verifica se o tipo de movimento é Deposito
    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_DEPOSITO_DINHEIRO Then gError 103804

    'preenche a tela
    Sequencial.Text = objMovimentoCaixa.lSequencial
    Data.PromptInclude = False
    Data.Text = Format(objMovimentoCaixa.dtDataMovimento, "dd/mm/yy")
    Data.PromptInclude = True
    Valor.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
    Historico.Text = objMovimentoCaixa.sHistorico

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 103805

    Traz_Deposito_Tela = SUCESSO

    Exit Function

Erro_Traz_Deposito_Tela:

    Traz_Deposito_Tela = gErr

    Select Case gErr

        Case 103802, 103803, 103805

        Case 103804
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO_DINHEIRO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158879)

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

    'preenche o campo data com a data atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    objTipoMeioPagtoLoja.iTipo = MEIO_PAGAMENTO_DINHEIRO
    objTipoMeioPagtoLoja.iFilialEmpresa = giFilialEmpresa
    
    'consulda o saldo
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTipoMeioPagtoLoja)
    If lErro <> SUCESSO Then gError 103800

    'preenche o campo valor com o saldo
    SaldoDinheiro.Caption = Format(objTipoMeioPagtoLoja.dSaldo, "STANDARD")

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103800

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158880)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Depósito Caixa"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "DepositoCaixa"
    
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

