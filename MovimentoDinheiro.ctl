VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl MovimentoDinheiro 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   7575
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5160
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   278
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "MovimentoDinheiro.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "F5 - Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "MovimentoDinheiro.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "F6 - Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "MovimentoDinheiro.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "F7 - Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "MovimentoDinheiro.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameIdentificacao 
      Caption         =   "Identificação"
      Height          =   870
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   4335
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1035
         TabIndex        =   1
         Top             =   345
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Mask            =   "#######"
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1860
         Picture         =   "MovimentoDinheiro.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton BotaoTrazer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2535
         Picture         =   "MovimentoDinheiro.ctx":0A7E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "F4 - Exibe na tela o movimento com o código informado."
         Top             =   200
         Width           =   1440
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.Frame FrameMovimentoDinheiro 
      Caption         =   "Movimento em Dinheiro"
      Height          =   3465
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   7095
      Begin VB.Frame FrameSuprimentoSangria 
         Caption         =   "Suprimento / Sangria"
         Height          =   975
         Left            =   600
         TabIndex        =   11
         Top             =   2280
         Width           =   5535
         Begin VB.OptionButton Suprimento 
            Caption         =   "Suprimento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   960
            TabIndex        =   5
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Sangria 
            Caption         =   "Sangria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3465
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin MSMask.MaskEdBox ValorMovimento 
         Height          =   645
         Left            =   2280
         TabIndex        =   4
         Top             =   1380
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1138
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Valor 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   600
         TabIndex        =   14
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label SaldoDinheiro 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2295
         TabIndex        =   12
         Top             =   360
         Width           =   3765
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   195
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MovimentoDinheiro.ctx":3748
   End
End
Attribute VB_Name = "MovimentoDinheiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Constantes Globais

Const TAMANHO_BUFFER_SEQUENCIAL = 10
Const NOME_ARQUIVO_CAIXA = "CAIXACONFIG.ini"
Const TRANSACAO_CAIXA_CONFIG = "Transacao"
Const CONSTANTE_ERRO = -1

Dim gcolImfCompl As New Collection

Dim iAlterado As Integer
Dim glProxNumAuto As Integer
Dim gsCodigo As String

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Sangria/Suprimento Dinheiro"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MovimentoDinheiro"

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

Private Sub Codigo_Validate(Cancel As Boolean)

Dim dSaldoEmDinheiro As Double

Dim lErro As Long

    If gsCodigo <> Codigo.Text Then
        
        gsCodigo = Codigo.Text
        
        lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
        If lErro <> SUCESSO Then
        
            SaldoDinheiro.Caption = ""
            
        Else
        
            'Joga o Saldo no caption do controle saldoDinheiro
            
            '??? 24/08/2016 SaldoDinheiro.Caption = Format(gdSaldoDinheiro, "standard")
            SaldoDinheiro.Caption = Format(dSaldoEmDinheiro, "standard")
        
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

'Tela Iniciada no Dia 1 de Agosto de 2002 por Sergio Ricardo

Public Sub Form_Load()
'Inicialização da Tela

Dim lErro As Long
Dim objSangria As Object
Dim objSuprimento As Object
    
Dim dSaldoEmDinheiro As Double

On Error GoTo Erro_Form_Load
    
    Set objSangria = Sangria
    Set objSuprimento = Suprimento
    
    Call CF_ECF("MovimentoDinheiro_Form_Load", objSangria, objSuprimento)
    
    'Exibe o Saldo do Caixa
    '??? 24/08/2016 SaldoDinheiro.Caption = Format(gdSaldoDinheiro, "Standard")
            
    lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    SaldoDinheiro.Caption = Format(dSaldoEmDinheiro, "Standard")
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
        
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163002)

    End Select

    Exit Sub
    
End Sub

Private Sub ValorMovimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorMovimento_Validate(Cancel As Boolean)
'Função que Valida se o Valor digitado é válido

Dim lErro As Long

On Error GoTo Erro_ValorMovimento_Validate

    'Verifica se Não foi digitado nada no Campo Referente ao Valor do Suprimento/ Sangria, se não sai da Função
    If Len(Trim(ValorMovimento.Text)) = 0 Then Exit Sub

    'Função que valida se o valor é Positivo
    lErro = Valor_Positivo_Critica(ValorMovimento.Text)
    If lErro <> SUCESSO Then gError 107680

    Exit Sub

Erro_ValorMovimento_Validate:

    Cancel = True

    Select Case gErr

        Case 107680
            ' Erro Tratado Dentro da Função que Foi Chamada
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163003)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
'Botão que Fecha o Form

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Fecha a Tela
    Unload Me

    Exit Sub
    
Erro_BotaoFechar_Click:
        
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163004)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()
'Botão que Chama a Função que Grava no Arquivo de MovimentoCaixa

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207975

    'Verifica se já foi executa a redução z para a data de hoje
    If gdtUltimaReducao = Date Then gError 111312
    
    '****** Alteração incluida para verificar se a sangria é maior que o valor que existe no caixa *******************
    
    If StrParaDbl(ValorMovimento.Text) > StrParaDbl(SaldoDinheiro.Caption) And Sangria.Value = True Then gError 111430

    '*********************Sergio dia 31/10/2002 **********************************************************************
    
    'Função que Grava o Movimento de Dinheiro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 107826
    
    'Função que Limpa Tela de Suprimento/Sangria Dinheiro
    Call Limpa_Tela(Me)
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 107826, 207975
    
        Case 111312
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
            
        Case 111430
            Call Rotina_ErroECF(vbOKOnly, ERRO_SANGRIA_MAIOR, gErr, Format(ValorMovimento.Text, "Standard"), Format(SaldoDinheiro.Caption, "Standard"))
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163005)
                
    End Select
    
    Exit Sub
    
End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim colMovimentosCaixa As New Collection
Dim iTipoMovimento As Integer
Dim vbMsgRes As VbMsgBoxResult

Dim dSaldoEmDinheiro As Double

On Error GoTo Erro_Gravar_Registro

    'Verifica se o Campo Relacionado ao Valor Está Preenchido se não Erro
    If Len(Trim(ValorMovimento.Text)) = 0 Then gError 107702
    
    'Verifica se o codigo está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107827
        
    'Verifica se a operação é de Sangria
    If Sangria.Value = True Then
        
        'Verifica se o valor da sangria é superior ao valor de dinheiro no Caixa se for Erro
        '??? 24/08/2016 If gdSaldoDinheiro + DELTA_VALORMONETARIO < StrParaDbl(ValorMovimento.Text) Then gError 107682
        
        lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If dSaldoEmDinheiro + DELTA_VALORMONETARIO < StrParaDbl(ValorMovimento.Text) Then gError 107682
        
    End If
    
    'verifica se já existe movimento com esse numero de Movimento
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, StrParaLong(Codigo.Text))
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107828
    
    'Se encontrou é uma alteração
    If colMovimentosCaixa.Count > 0 Then
    
        Set objMovimentoCaixa = colMovimentosCaixa(1)
    
        If (objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SANGRIA_DINHEIRO) And _
           (objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO) Then gError 86286
           
        'Envia aviso perguntando se deseja atualizar o movimemtos
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ALTERACAO_MOVIMENTOCAIXA, Codigo.Text)

        'Se a Reposta for Negativa
        If vbMsgRes = vbNo Then gError 107829

        'Procura na coleção que veio carregada
        For Each objMovimentoCaixa In colMovimentosCaixa
            'se o tipo de movimento for de sangria de dinheiro
            If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then
            
                iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_DINHEIRO
                
            'senão
            Else
            
                iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SUPRIMENTO_DINHEIRO
                
            End If
            
        Next
        
        'Função que Faz a Alteração na Sangria de Boleto Previamente Executada, adciona o iTipoMovimento
        lErro = MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa, iTipoMovimento)
        If lErro <> SUCESSO Then gError 107830
        
    End If
    'Move os Dados da Tela para memória
    lErro = Move_Dados_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107856
    
    'Grava o Movimento no Arquivo Caixa
    lErro = Caixa_Grava_MovimentoDinheiro(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107831
    
    'Chama a Função que Atauliza a Memoria
    lErro = MovimentoDinheiro_Atualiza_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107834
    
    'Chama a Função que Limpa a Tela e Atauliza a label da quantia que Esta no Caixa
    Call Limpa_Tela(Me)
    
    Suprimento.Value = True

    'Exibe o Saldo no Caixa
    '??? 24/08/2016 SaldoDinheiro.Caption = Format(gdSaldoDinheiro, "Standard")
    
    lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    SaldoDinheiro.Caption = Format(dSaldoEmDinheiro, "Standard")
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:
        
    Gravar_Registro = gErr
        
    Select Case gErr
    
        Case 86286
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_DINHEIRO, gErr, StrParaDbl(Codigo.Text))

        Case 107682
            Call Rotina_ErroECF(vbOKOnly, ERRO_SALDO_INSUFICIENTE_SANGRIA, gErr, dSaldoEmDinheiro, giCodCaixa, ValorMovimento.Text)
        
        Case 107702
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
        
        Case 107827
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)
        
        Case 107828, 107829, 107830, 107831, 107834, 107856, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163006)

    End Select

    Exit Function

End Function

Function Caixa_Grava_MovimentoDinheiro(colMovimentosCaixa As Collection) As Long
'Função que Grava os Movimentos de Caixa Relacionados a Dinheiro

Dim lErro As Long
Dim objMovimentoCaixa As ClassMovimentoCaixa
Dim lSequencial As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objOperador As New ClassOperador
Dim iIndice As Integer
Dim colRegistro As New Collection
Dim objMovCx As ClassMovimentoCaixa
Dim sNomeArq As String
Dim lTamanho As Long
Dim sRetorno As String
Dim sArquivo As String

Dim sMensagem As String

On Error GoTo Erro_Caixa_Grava_MovimentoDinheiro
        
    'Verifica o Status da Sessão de Caixa
    If giStatusSessao = SESSAO_ENCERRADA Then

        'Envia aviso perguntando se de seja Abrir sessão
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_DESEJA_ABRIR_SESSAO, giCodCaixa)

        If vbMsgRes = vbNo Then gError 107835

        'Função que Executa Abertura na Sessão
        lErro = CF_ECF("Sessao_Executa_Abertura")
        If lErro <> SUCESSO Then gError 107837

    End If

    'Se for Necessário a Altorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 107838

    End If

    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosCC", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)
    
    'Se não encontrou
    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 127097
    
    If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
    
    sArquivo = sRetorno & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC
    
    'Abre o arquivo de retorno
    Open sArquivo For Input Lock Read Write As #10

    'Função que Abre a Transação de Caixa, Identificador dentro do Caixa para um determinado MOVTO
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 107839
        
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
        
    'Obtém a ultima transacao transferida
    Call GetPrivateProfileString(APLICACAO_DADOS, "UltimaTransacaoTransf", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)
        
    For Each objMovimentoCaixa In colMovimentosCaixa

        'se o numero da ultima transacao transferida ultrapassar o numero da transacao do movimento de caixa
        If objMovimentoCaixa.lSequencial <> 0 And StrParaLong(sRetorno) > objMovimentoCaixa.lSequencial Then gError 133845

        'Caso nao precise de autorizacao do gerente nesta transacao ==> objOperador.iCodigo vai estar zerado
        objMovimentoCaixa.iGerente = objOperador.iCodigo
        
        lErro = Caixa_Grava_MovCx(objMovimentoCaixa, lSequencial)
        If lErro <> SUCESSO Then gError 105709
        
    Next

    lSequencial = lSequencial - 1
    
    'Fecha a Transação
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 107842

    Close #10
    
    Caixa_Grava_MovimentoDinheiro = SUCESSO

    Exit Function

Erro_Caixa_Grava_MovimentoDinheiro:

    Close #10

    Caixa_Grava_MovimentoDinheiro = gErr

    Select Case gErr

        Case 105709, 107839, 107842
            
        Case 107835, 107837, 107838
        
        Case 133845
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163007)

    End Select

    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
    
    Exit Function

End Function

Function Caixa_Grava_MovCx(objMovimentoCaixa As ClassMovimentoCaixa, lSequencial As Long) As Long
'grava cada movimento de caixa passado como parametro

Dim lErro As Long
Dim colRegistro As New Collection
Dim sMensagem As String
Dim objMovCx As ClassMovimentoCaixa
Dim objTela As Object

On Error GoTo Erro_Caixa_Grava_MovCx

    'para não ficar 3 movimentos com o mesmo Código(Numero de Movto) na Coleção gcolMovto
    If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_DINHEIRO Or objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SUPRIMENTO_DINHEIRO Then

        Set objMovCx = New ClassMovimentoCaixa
            
        lErro = CF_ECF("MovimentoCaixa_Copia", objMovimentoCaixa, objMovCx)
        If lErro <> SUCESSO Then gError 105703

    Else
    
        Set objMovCx = objMovimentoCaixa

    End If

    'Guarda o Sequencial no objmovimentoCaixa
    objMovCx.lSequencial = lSequencial

    lSequencial = lSequencial + 1

    'Guarda no objMovimentoCaixa os Dados que Serão Usados para a Geração do Movimento de Caixa
    lErro = CF_ECF("Move_DadosGlobais_Memoria", objMovCx)
    If lErro <> SUCESSO Then gError 107840

    'Funçao que Gera o Arquivo preparando para a gravação
    Call CF_ECF("MovimentoDinheiro_Gera_Log", colRegistro, objMovCx)

    'Função que Vai Gravar as Informações no Arquivo de Caixa
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
    If lErro <> SUCESSO Then gError 107841
    
    Set colRegistro = New Collection
    Set objTela = Me
    

    'Faz a sangria
    lErro = CF_ECF("Sangria_AFRAC", objMovCx.dValor, sMensagem, objMovCx.iTipo, objTela)
    If lErro <> SUCESSO Then gError 109806
    
    Caixa_Grava_MovCx = SUCESSO
    
    Exit Function

Erro_Caixa_Grava_MovCx:

    Caixa_Grava_MovCx = gErr

    Select Case gErr

        Case 105703, 107840, 107841, 109806, 109815

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163008)

    End Select
    
    Exit Function

End Function

Function Move_Dados_Memoria(colMovimentosCaixa As Collection) As Long
'Função que Move os dados para a memoria

Dim lErro As Long
Dim objMovimentosCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Move_Dados_Memoria

    'Guardo o codigo do movimento
    objMovimentosCaixa.lNumMovto = StrParaLong(Codigo.Text)
    
    objMovimentosCaixa.iParcelamento = COD_A_VISTA
    
    objMovimentosCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
    
    'Se a Sangria estiver selecionada
    If Sangria.Value = True Then
    
        objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO
        
    ElseIf Suprimento.Value = True Then
    
        objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO
        
    End If
    
'??? 24/08/2016     If gdSaldoDinheiro < StrParaDbl(ValorMovimento.Text) And Sangria.Value = True Then
'??? 24/08/2016         objMovimentosCaixa.dValor = gdSaldoDinheiro
'??? 24/08/2016     Else
    
        'Guarda o valor do movimento
        objMovimentosCaixa.dValor = StrParaDbl(ValorMovimento.Text)
'??? 24/08/2016     End If
    
    'Adciona na Coleção
    colMovimentosCaixa.Add objMovimentosCaixa
    
    Move_Dados_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_Memoria:

    Move_Dados_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163009)

    End Select

    Exit Function

End Function

Function MovimentoDinheiro_Atualiza_Memoria(colMovimentosCaixa As Collection) As Long
'Função que Atualiza a Memoria

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_MovimentoDinheiro_Atualiza_Memoria

    For Each objMovimentoCaixa In colMovimentosCaixa

        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_DINHEIRO Or objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SUPRIMENTO_DINHEIRO Then

            'Função que Retira de memória os Movimentos Excluidos
            lErro = MovimentoCaixa_Exclui_Memoria(objMovimentoCaixa)
            If lErro <> SUCESSO Then gError 107843

'??? 24/08/2016             If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_DINHEIRO Then
'??? 24/08/2016
'??? 24/08/2016                 'Atualiza o Saldo Global
'??? 24/08/2016                 gdSaldoDinheiro = gdSaldoDinheiro + objMovimentoCaixa.dValor
'??? 24/08/2016
'??? 24/08/2016             Else
'??? 24/08/2016
'??? 24/08/2016                 gdSaldoDinheiro = gdSaldoDinheiro - objMovimentoCaixa.dValor
'??? 24/08/2016
'??? 24/08/2016             End If
            
        Else
        
            'Adcionar a Coleção Global o objMovimento Caixa
            gcolMovimentosCaixa.Add objMovimentoCaixa
            
'??? 24/08/2016             'Atualiza a Variável Global
'??? 24/08/2016             If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then
'??? 24/08/2016
'??? 24/08/2016                 gdSaldoDinheiro = gdSaldoDinheiro - objMovimentoCaixa.dValor
'??? 24/08/2016
'??? 24/08/2016             ElseIf objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then
'??? 24/08/2016
'??? 24/08/2016                 gdSaldoDinheiro = gdSaldoDinheiro + objMovimentoCaixa.dValor
'??? 24/08/2016
'??? 24/08/2016             End If
        
        End If

    Next

    MovimentoDinheiro_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoDinheiro_Atualiza_Memoria:

    MovimentoDinheiro_Atualiza_Memoria = gErr

    Select Case gErr

        Case 107843

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163010)

    End Select

    Exit Function

End Function

Function MovimentoCaixa_Exclui_Memoria(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Exclui da Memória os Movimentos que Foram Alterados

Dim lErro As Long
Dim objMovimentoCaixaAux As New ClassMovimentoCaixa
Dim iIndice As Integer

On Error GoTo Erro_MovimentoCaixa_Exclui_Memoria

    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1

        Set objMovimentoCaixaAux = gcolMovimentosCaixa.Item(iIndice)

        'Verifica se o movimento é o mesmo
        If objMovimentoCaixa.lNumMovto = objMovimentoCaixaAux.lNumMovto And objMovimentoCaixa.lSequencial = objMovimentoCaixaAux.lSequencial Then

            'Exclui o movimento da Coleção Global de MovimentosCaixa
            gcolMovimentosCaixa.Remove (iIndice)

            'Sai do Loop
            Exit For

        End If

    Next

    MovimentoCaixa_Exclui_Memoria = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Exclui_Memoria:

    MovimentoCaixa_Exclui_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163011)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoProxNum_Click()
'Botão que Gera um Próximo Numero para Movto

Dim lErro As Long
Dim lNumero As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Função que Gera o Próximo Código para a Tela de Sangria de Boletos
    lErro = CF_ECF("Caixa_Obtem_NumAutomatico", lNumero)
    If lErro <> SUCESSO Then gError 107820

    'Exibir o Numero na Tela
    Codigo.Text = lNumero

    gsCodigo = Codigo.Text

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 107820

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163012)

    End Select

    Exit Sub

End Sub


Private Sub BotaoTrazer_Click()
'Função que chama a função que preenche o grid

Dim lErro As Long
Dim sCodigo As String

Dim dSaldoEmDinheiro As Double

On Error GoTo Erro_botaoTrazer_click

    'Verifica se o código não está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107823

    'Joga o Saldo no caption do controle saldoDinheiro
    '??? 24/08/2016 SaldoDinheiro.Caption = Format(gdSaldoDinheiro, "standard")
    
    lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    SaldoDinheiro.Caption = Format(dSaldoEmDinheiro, "standard")
    
    'Chama a função o Movimento de Sangria de Dinheiro
    lErro = Traz_MovimentoDinheiro_Tela(StrParaLong(Codigo.Text))
    If lErro <> SUCESSO Then gError 107824

    'Anula a Alteração
    iAlterado = 0
    
    Exit Sub

Erro_botaoTrazer_click:

    Select Case gErr

        Case 107823
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 107824, ERRO_SEM_MENSAGEM
            'Erro tradado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163013)

    End Select

    Exit Sub

End Sub

Private Sub SaldoDinheiro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Suprimento_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sangria_Click()

       iAlterado = REGISTRO_ALTERADO

End Sub

Function Traz_MovimentoDinheiro_Tela(lNumero As Long) As Long
'Função que Traz o Movto de Dinheiro para a Tela

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim objMovimentoCaixa As New ClassMovimentoCaixa

Dim dSaldoEmDinheiro As Double

On Error GoTo Erro_Traz_MovimentoDinheiro_Tela

    'Função que Lê os Movimentos Carregado
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107825
    
    If lErro = 107850 Then gError 107852
    
    'Varre a Coleção, a procura do MOVTO de Código passado
    For Each objMovimentoCaixa In colMovimentosCaixa
    
        'Verifica se o movimento é do tipo Movimento Sangria Dinheiro ou Movimento Suprimento Dinheiro
        If objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SANGRIA_DINHEIRO And objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then gError 105739
            
        'Atribui o valor do Movimento a variável global
        '??? 24/08/2016 SaldoDinheiro.Caption = Format(IIf(objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO, objMovimentoCaixa.dValor, -objMovimentoCaixa.dValor) + gdSaldoDinheiro, "standard")
            
        lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        SaldoDinheiro.Caption = Format(IIf(objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO, objMovimentoCaixa.dValor, -objMovimentoCaixa.dValor) + dSaldoEmDinheiro, "standard")
            
        'Exibir os valores na Tela
        ValorMovimento.Text = Format(objMovimentoCaixa.dValor, "standard")
        
        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then
            Sangria.Value = True
        Else
            Suprimento.Value = True
        End If
            
    Next
    
    Traz_MovimentoDinheiro_Tela = SUCESSO
    
    Exit Function

Erro_Traz_MovimentoDinheiro_Tela:

    Traz_MovimentoDinheiro_Tela = gErr
    
    Select Case gErr

        Case 105739
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_DINHEIRO, gErr, lNumero)

        Case 107825, ERRO_SEM_MENSAGEM

        Case 107852
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163014)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim lNumero As Long
Dim iTipoMovimento As Integer
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim dValorAtualizar As Double

Dim dSaldoEmDinheiro As Double

On Error GoTo Erro_BotaoExcluir_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207976

    'Verifica se já foi executa a redução z para a data de hoje
    If gdtUltimaReducao = Date Then gError 111313

    'Verifica se o Codigo não foi Preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107844

    lNumero = StrParaLong(Codigo.Text)

    'Verifica os Movimentos de Caixa para o Código em Questão
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107851
    
    'Se não encontrou
    If lErro = 107850 Then gError 107845
       
    'Pergunta se deseja Realmente Excluir o Movimento
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_EXCLUSAO_MOVIMENTOCAIXA, Codigo.Text)

    If vbMsgRes = vbNo Then gError 107846

    For Each objMovimentoCaixa In colMovimentosCaixa

        If (objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SANGRIA_DINHEIRO) And _
           (objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO) Then gError 86286
           
        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_DINHEIRO Then
        
            iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_DINHEIRO

        ElseIf objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SUPRIMENTO_DINHEIRO Then
        
            iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SUPRIMENTO_DINHEIRO
        
        End If
        
    Next
    
    'Prepara os Movimentos para a Exclusão
    lErro = MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa, iTipoMovimento)
    If lErro <> SUCESSO Then gError 107847

    'Função que Grava a Exclusão de Boletos
    lErro = Caixa_Grava_MovimentoDinheiro(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107848

    'Atualiza os Dados na Memória
    lErro = MovimentoDinheiro_Atualiza_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107849
    
    'Função Que Limpa a Tela
    Call Limpa_Tela(Me)

    'Joga o Saldo na Tela
    '??? 24/08/2016 SaldoDinheiro.Caption = Format(gdSaldoDinheiro, "standard")
    
    lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    SaldoDinheiro.Caption = Format(dSaldoEmDinheiro, "standard")
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 86286
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_DINHEIRO, gErr, Codigo.Text)
        
        Case 86289
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
        
        Case 107844
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 107845
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)

        Case 107846, 107847, 107848, 107849, 107851, 207976, ERRO_SEM_MENSAGEM

        Case 111312
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163015)

    End Select

    Exit Sub

End Sub

Function MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa As Collection, Optional iTipoMovimento As Integer) As Long
'Função que Para Cada obj da Coleção adciona a esse movimento dizando q foi alterado o valor da Sangria

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_MovimentoCaixa_Prepara_Exclusao

    For Each objMovimentoCaixa In colMovimentosCaixa

        'Adciona o Tipo de Movimento a Coleção de Movimentos
        objMovimentoCaixa.iTipo = iTipoMovimento

    Next

    MovimentoCaixa_Prepara_Exclusao = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Prepara_Exclusao:

    MovimentoCaixa_Prepara_Exclusao = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163016)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'Função que Limpa a Tela

Dim lErro As Long
Dim dSaldoEmDinheiro As Double

    Call Limpa_Tela(Me)
    
    gsCodigo = ""
    
    'Joga o Saldo no caption do controle saldoDinheiro
    '??? 24/08/2016 SaldoDinheiro.Caption = Format(gdSaldoDinheiro, "standard")

    lErro = CF_ECF("SaldoEmDinheiro_Le", dSaldoEmDinheiro)
    If lErro <> SUCESSO Then
        SaldoDinheiro.Caption = ""
    Else
        SaldoDinheiro.Caption = Format(dSaldoEmDinheiro, "standard")
    End If
    
    'Por default deixa marcada sangria
    Sangria.Value = True
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Função que Incrementa o Código Atravez da Tecla F2
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case KEYCODE_PROXIMO_NUMERO
            
            'Função que Incrementa o Código( Ultimo Movto + 1)
            BotaoProxNum.SetFocus
            If Not TrocaFoco(Me, BotaoProxNum) Then Exit Sub
            Call BotaoProxNum_Click

        Case KEYCODE_BROWSER

            Call LabelCodigo_Click

        Case vbKeyF4
            If Not TrocaFoco(Me, BotaoTrazer) Then Exit Sub
            Call BotaoTrazer_Click

        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click
            
        Case vbKeyF6
            If Not TrocaFoco(Me, BotaoExcluir) Then Exit Sub
            Call BotaoExcluir_Click
            
        Case vbKeyF7
            If Not TrocaFoco(Me, BotaoLimpar) Then Exit Sub
            Call BotaoLimpar_Click
            
        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click
            

    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163017)

    End Select

    Exit Sub

End Sub

Sub CodMovimentoDinheiro_Validate(Cancel As Boolean)
'Função que Verifica se o Código Passado como parâmetro Existe na Coleção Globa de MOVTOCAIXA

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_CodMovimentoDinheiro_Validate

    'Verifica se existe movimento com o código passado
    For Each objMovimentoCaixa In gcolMovimentosCaixa
    
        If objMovimentoCaixa.lNumMovto = StrParaLong(Codigo.Text) Then
            
            'Função que traz o MovimentoBoleto para a Tela
            lErro = Traz_MovimentoDinheiro_Tela(StrParaLong(Codigo.Text))
            If lErro <> SUCESSO Then gError 107819
            Exit For
        
        End If
     
    
    Next
    
    'Anula a Alteração
    iAlterado = 0
    
    Exit Sub
    
Erro_CodMovimentoDinheiro_Validate:
    
    Select Case gErr

        Case 107819

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163018)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objMovimentoCaixa As New ClassMovimentoCaixa
    
    'Chama tela de MovimentoBoletoLista
    Call Chama_TelaECF_Modal("MovimentoDinheiroLista", objMovimentoCaixa)
    
    If Not (objMovimentoCaixa Is Nothing) Then
        'Verifica se o Codvendedor está preenchido e joga na coleção
        If objMovimentoCaixa.lNumMovto <> 0 Then
            Codigo.Text = objMovimentoCaixa.lNumMovto
            gsCodigo = Codigo.Text
            Call CodMovimentoDinheiro_Validate(False)
            
        End If
    End If
    
    Exit Sub

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set gcolImfCompl = Nothing
End Sub
