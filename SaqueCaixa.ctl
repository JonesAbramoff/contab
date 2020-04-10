VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl SaqueCaixa 
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   7500
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa Central"
      Height          =   810
      Left            =   75
      TabIndex        =   18
      Top             =   3420
      Width           =   7275
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
         TabIndex        =   20
         Top             =   570
         Width           =   1050
      End
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
         Top             =   270
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.Label SaldoDinheiro 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3330
         TabIndex        =   22
         Top             =   330
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
         Left            =   1650
         TabIndex        =   21
         Top             =   375
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Principais"
      Height          =   1665
      Left            =   90
      TabIndex        =   14
      Top             =   720
      Width           =   7275
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2670
         Picture         =   "SaqueCaixa.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   255
         Width           =   300
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   5250
         TabIndex        =   3
         Top             =   240
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
         Left            =   6405
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   1635
         TabIndex        =   5
         Top             =   705
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
         Left            =   1620
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Favorecido 
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         Top             =   1155
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   50
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
         Left            =   525
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   285
         Width           =   1020
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
         Left            =   1035
         TabIndex        =   17
         Top             =   750
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
         Index           =   1
         Left            =   4695
         TabIndex        =   16
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Favorecido:"
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
         Index           =   0
         Left            =   525
         TabIndex        =   15
         Top             =   1200
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5220
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "SaqueCaixa.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "SaqueCaixa.ctx":0274
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "SaqueCaixa.ctx":03F2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "SaqueCaixa.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Complemento"
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   2505
      Width           =   7275
      Begin MSMask.MaskEdBox Historico 
         Height          =   315
         Left            =   1710
         TabIndex        =   7
         Top             =   285
         Width           =   4935
         _ExtentX        =   8705
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
         Left            =   810
         TabIndex        =   12
         Top             =   345
         Width           =   825
      End
   End
End
Attribute VB_Name = "SaqueCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Buttons CaixaCentral e CaixaGeral são meramente informativos. Não
'podem ser editados.

Private WithEvents objEventoSaque As AdmEvento
Attribute objEventoSaque.VB_VarHelpID = -1

'variáveis globais
Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Function Mover_Dados_Saque_CCMovDia(objMovimentoCaixa As ClassMovimentoCaixa, objCCMovDia As ClassCCMovDia) As Long
'Função que recebe um objMovimentoCaixa com todos os dados preenchidos e preenche o objCCmovDia para sua atualalização
    
On Error GoTo Erro_Mover_Dados_Saque_CCMovDia

    'preenche os dados do objCCMovDia com os dados do movimento de caixa passado por parâmetro
    objCCMovDia.dDeb = objMovimentoCaixa.dValor
    objCCMovDia.dtData = objMovimentoCaixa.dtDataMovimento
    objCCMovDia.iCodCaixa = objMovimentoCaixa.iCaixa
    objCCMovDia.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto
    objCCMovDia.iFilialEmpresa = objMovimentoCaixa.iFilialEmpresa
    objCCMovDia.iParcelamento = objMovimentoCaixa.iParcelamento
    objCCMovDia.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_DINHEIRO

    Mover_Dados_Saque_CCMovDia = SUCESSO

    Exit Function

Erro_Mover_Dados_Saque_CCMovDia:

    Mover_Dados_Saque_CCMovDia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174335)

    End Select

    Exit Function

End Function

Private Sub objEventoSaque_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovimentoCaixa As ClassMovimentoCaixa

On Error GoTo Erro_objEventoSaque_evSelecao

    Set objMovimentoCaixa = obj1
    
    'tenta ler do bd
    lErro = CF("MovimentosCaixa_Le", objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103744 Then gError 103788
    
    'se não encontrou-> erro
    If lErro = 103744 Then gError 103789
    
    'se o movimento encontrado no bd não for do tipo saque, erro
    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_SAQUE Then gError 103790
    
    'mover os dados para a tela
    Sequencial.Text = objMovimentoCaixa.lSequencial
    
    Data.PromptInclude = False
    Data.Text = Format(objMovimentoCaixa.dtDataMovimento, "dd/mm/yy")
    Data.PromptInclude = True
    
    Valor.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
    Favorecido.Text = objMovimentoCaixa.sFavorecido
    Historico.Text = objMovimentoCaixa.sHistorico
    
    iAlterado = 0
    
    'mostrar a tela
    Me.Show

    Exit Sub

Erro_objEventoSaque_evSelecao:

    Select Case gErr
    
        Case 103788
        
        Case 103789
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOSCAIXA_NAOENCONTRADO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)
            
        Case 103790
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_SAQUE", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174336)

    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_Saque()

Dim lErro As Long
Dim dValor As Double
Dim objTipoMeioPagtoLoja As New ClassTMPLojaFilial

On Error GoTo Erro_Limpa_Tela_Saque

    'limpa os campos da tela
    Call Limpa_Tela(Me)
    
    'preenche o tipo meio pagamento comm pagamento dinheiro
    objTipoMeioPagtoLoja.iTipo = MEIO_PAGAMENTO_DINHEIRO
    objTipoMeioPagtoLoja.iFilialEmpresa = giFilialEmpresa
    
    'le o saldo atualizado da tabela
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTipoMeioPagtoLoja)
    If lErro <> SUCESSO Then gError 103787
    
    'preenche o campo valor com o saldo
    SaldoDinheiro.Caption = Format(objTipoMeioPagtoLoja.dSaldo, "STANDARD")
    
    'preenche o campo data com a data atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    Exit Sub

Erro_Limpa_Tela_Saque:

    Select Case gErr
    
        Case 103787
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174337)
            
    End Select
    
    Exit Sub

End Sub

Public Sub form_unload(Cancel As Integer)

On Error GoTo Erro_Form_Unload
    
    'libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)
    Set objEventoSaque = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174338)
            
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

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'testa se houve alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 103758

    'limpa os campos que se fizerem necessários
    Call Limpa_Tela_Saque

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 103759

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 103758, 103759

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174339)

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
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 126045

    'verifica se sequencial está preenchido
    If Len(Trim(Sequencial.Text)) = 0 Then gError 103735

    objMovimentoCaixa.iCaixa = CODIGO_CAIXA_CENTRAL
    objMovimentoCaixa.lSequencial = StrParaLong(Sequencial.Text)

    lErro = CF("MovimentosCaixa_Le", objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103744 Then gError 103736

    If lErro = 103744 Then gError 103737

    'se o movimento não for do tipo saque--> erro
    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_SAQUE Then gError 103738

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_MOVIMENTOSCAIXA_EXCLUSAO", STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)
    
    If vbMsgRes = vbYes Then

        objLog.iOperacao = EXCLUSAO_SAQUE_CAIXA_CENTRAL

        Call Mover_Dados_Saque_Log_Exc(objMovimentoCaixa, objLog)
    
        lErro = Mover_Dados_Saque_CCMovDia(objMovimentoCaixa, objCCMovDia)
        If lErro <> SUCESSO Then gError 103739

        lErro = CF("MovimentosCaixa_Exclui", objMovimentoCaixa, objLog, objCCMovDia)
        If lErro <> SUCESSO Then gError 103740

        Call Limpa_Tela_Saque

        iAlterado = 0

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 103735
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)

        Case 103736, 103739, 103740

        Case 103737
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOSCAIXA_NAOENCONTRADO", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case 103738
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_SAQUE", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case 126045
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NAO_PERMITIDA_BACKOFFICE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174340)

    End Select

    Exit Sub

End Sub

Private Sub Mover_Dados_Saque_Log_Exc(objMovimentoCaixa As ClassMovimentoCaixa, objLog As ClassLog)
'Função que recebe um objMovimentoCaixa com todos os dados preenchidos e preenche a strin de um
'objLog para prepará-lo para a gravação

On Error GoTo Erro_Mover_Dados_Saque_Log_Exc

    'preenche os dados do log com cada atributo do objMovimentoCaixa separado por vbkeyscape
    With objMovimentoCaixa
        objLog.sLog = CStr(.iFilialEmpresa) & Chr(vbKeyEscape) & _
                      CStr(.iCaixa) & Chr(vbKeyEscape) & _
                      CStr(.lSequencial) & Chr(vbKeyEscape) & _
                      Chr(vbKeyEnd)

    End With

    Exit Sub

Erro_Mover_Dados_Saque_Log_Exc:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174341)

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
    objMovimentoCaixa.iTipo = MOVIMENTO_CAIXA_SAQUE
    objMovimentoCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
    objMovimentoCaixa.iParcelamento = PARCELAMENTO_AVISTA
    objMovimentoCaixa.dtDataMovimento = StrParaDate(Data.Text)
    objMovimentoCaixa.dValor = StrParaDbl(Valor.Text)
    objMovimentoCaixa.sFavorecido = Trim(Favorecido.Text)
    objMovimentoCaixa.sHistorico = Trim(Historico.Text)
    objMovimentoCaixa.dHora = CDbl(Time)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174342)

    End Select

    Exit Function

End Function

Private Function Traz_Saque_Tela(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que preenche a tela com os dados selecionados no browser

Dim lErro As Long

On Error GoTo Erro_Traz_Saque_Tela

    'busca o movimento de caixa
    lErro = CF("MovimentosCaixa_Le_NumIntDoc", objMovimentoCaixa)
    If lErro <> SUCESSO And lErro <> 103677 Then gError 103670

    If lErro = 103677 Then gError 103671

    'verifica se o tipo de movimento é saque
    If objMovimentoCaixa.iTipo <> MOVIMENTO_CAIXA_SAQUE Then gError 103672

    'preenche a tela
    Sequencial.Text = objMovimentoCaixa.lSequencial
    Data.PromptInclude = False
    Data.Text = Format(objMovimentoCaixa.dtDataMovimento, "dd/mm/yy")
    Data.PromptInclude = True
    Valor.Text = Format(objMovimentoCaixa.dValor, "STANDARD")
    Favorecido.Text = objMovimentoCaixa.sFavorecido
    Historico.Text = objMovimentoCaixa.sHistorico

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 103673

    Traz_Saque_Tela = SUCESSO

    Exit Function

Erro_Traz_Saque_Tela:

    Traz_Saque_Tela = gErr

    Select Case gErr

        Case 103670, 103671, 103673

        Case 103672
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_SAQUE", gErr, STRING_CAIXA_CENTRAL, objMovimentoCaixa.lSequencial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174343)

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
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 126044
    
    'verifica se o seqüencial está preenchido
    If Len(Trim(Sequencial.Text)) = 0 Then gError 103665

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 103666

    'carrega o objCaixa com os campos chave
    objCaixa.iFilialEmpresa = giFilialEmpresa
    objCaixa.iCodigo = CODIGO_CAIXA_CENTRAL

    'verifica se o valor está preenchido
    If Len(Trim(Valor.Text)) = 0 Then gError 103667

    'preenche o obj com os dados da tela
    lErro = Move_Tela_Memoria(objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103668

    'Pergunta se deseja alterar
    lErro = Trata_Alteracao(objMovimentoCaixa, objMovimentoCaixa.iCaixa, objMovimentoCaixa.lSequencial)
    If lErro <> SUCESSO Then gError 107049
    
    'grava o saque
    lErro = CF("Movimentos_Caixa_Grava_Saque", objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103669
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 103665
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", gErr)

        Case 103666
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 103667
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_SAQUE_NAO_PREENCHIDO", gErr)

        Case 103668, 103669, 107049

        Case 126044
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_NAO_PERMITIDA_BACKOFFICE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174344)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'grava o registro
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 103664

    'limpa a tela
    Call Limpa_Tela_Saque

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 103664

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174345)

    End Select

    Exit Sub

End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Favorecido_Change()

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
        If lErro <> SUCESSO Then gError 103661

    End If
    
    Cancel = False

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 103661

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174346)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174347)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'aumenta a data de um dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 103660

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 103660

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174348)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'diminui a data de um dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 103659

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 103659

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174349)

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
    If lErro <> SUCESSO Then gError 103658
    
    Cancel = False

    Exit Sub

Erro_Data_Validate:

    Cancel = True
    
    Select Case gErr

        Case 103658

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174350)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174351)

    End Select

Exit Sub

End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)

On Error GoTo Erro_Sequencial_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(Sequencial.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(Sequencial.Text) Then gError 103656

        'Verifica se codigo é menor que um
        If StrParaLong(Sequencial.Text) < 1 Then gError 103657

    End If

    Cancel = False

    Exit Sub

Erro_Sequencial_Validate:

    Cancel = True

    Select Case gErr

        Case 103656, 103657
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_INVALIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174352)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174353)

    End Select

Exit Sub

End Sub

Private Sub Sequencial_Change()

    iAlterado = REGISTRO_ALTERADO

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
    If lErro <> SUCESSO Then gError 103655

    Sequencial.Text = lSequencial

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 103655

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174354)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Tela_Preenche

    'preenche o objmovimentocaixa com a colecao de valores
    objMovimentoCaixa.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objMovimentoCaixa.lNumMovto = colCampoValor.Item("NumMovto").vValor

    'traz os dados do saque para a tela
    lErro = Traz_Saque_Tela(objMovimentoCaixa)
    If lErro <> SUCESSO Then gError 103654
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 103654
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174355)

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
    If lErro <> SUCESSO Then gError 103653

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
    colSelecao.Add "Tipo", OP_IGUAL, MOVIMENTO_CAIXA_SAQUE

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 103653

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174356)

    End Select

    Exit Function

End Function

Private Sub LabelSequencial_Click()

Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim colSelecao As New Collection

On Error GoTo Erro_LabelSequencial_Click

    'se o sequencial estiver preenchido
    If Len(Trim(Sequencial.Text)) <> 0 Then

        'preenche o atributo seqüencial do obj com o conteúdo do campo seqüencial
        objMovimentoCaixa.lSequencial = StrParaLong(Trim(Sequencial.Text))

    End If

    Call Chama_Tela("SaqueLojaLista", colSelecao, objMovimentoCaixa, objEventoSaque)

    Exit Sub

Erro_LabelSequencial_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174357)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objMovimentoCaixa As ClassMovimentoCaixa) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se há movimento de caixa
    If Not (objMovimentoCaixa Is Nothing) Then

        'traz os dados do saque para a tela
        lErro = Traz_Saque_Tela(objMovimentoCaixa)
        If lErro <> SUCESSO And lErro <> 103671 Then gError 103651

        'se retornou erro indicando que não está cadastrado
        If lErro = 103671 Then

            'limpa a tela
            Call Limpa_Tela_Saque

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

        Case 103651

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174358)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim dValor As Double
Dim objTipoMeioPagtoLoja As New ClassTMPLojaFilial

On Error GoTo Erro_Form_Load

    'seta o admEvento
    Set objEventoSaque = New AdmEvento

    'preenche o campo data com a data atual
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    objTipoMeioPagtoLoja.iTipo = MEIO_PAGAMENTO_DINHEIRO
    objTipoMeioPagtoLoja.iFilialEmpresa = giFilialEmpresa

    'consulda o saldo
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTipoMeioPagtoLoja)
    If lErro <> SUCESSO Then gError 103650

    'preenche o campo valor com o saldo
    SaldoDinheiro.Caption = Format(objTipoMeioPagtoLoja.dSaldo, "STANDARD")

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103650

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174359)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Saque Caixa"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "SaqueCaixa"

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
