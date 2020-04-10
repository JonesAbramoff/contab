VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ContratoPropaganda 
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   ScaleHeight     =   2850
   ScaleWidth      =   5670
   Begin VB.CommandButton BotaoContrato 
      Caption         =   "Consulta Contratos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3945
      TabIndex        =   8
      Top             =   2145
      Width           =   1575
   End
   Begin VB.Frame FrameData 
      Caption         =   "Período"
      Height          =   750
      Left            =   165
      TabIndex        =   10
      Top             =   705
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownPeriodoDe 
         Height          =   330
         Left            =   1980
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox PeriodoDe 
         Height          =   315
         Left            =   990
         TabIndex        =   0
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownPeriodoAte 
         Height          =   330
         Left            =   4575
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox PeriodoAte 
         Height          =   330
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPeriodoDe 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   585
         TabIndex        =   14
         Top             =   300
         Width           =   390
      End
      Begin VB.Label LabelPeriodoAte 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   3180
         TabIndex        =   13
         Top             =   285
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ContratoPropagandaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ContratoPropagandaOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ContratoPropagandaOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ContratoPropagandaOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   1155
      TabIndex        =   2
      Top             =   1650
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Percentual 
      Height          =   300
      Left            =   1155
      TabIndex        =   3
      Top             =   2160
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#0.#0\%"
      PromptChar      =   " "
   End
   Begin VB.Label LabelPercentual 
      Caption         =   "Percentual:"
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
      Left            =   75
      TabIndex        =   16
      Top             =   2205
      Width           =   960
   End
   Begin VB.Label LabelCliente 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
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
      Left            =   405
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   1695
      Width           =   660
   End
End
Attribute VB_Name = "ContratoPropaganda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

'Browser
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoBotaoContrato As AdmEvento
Attribute objEventoBotaoContrato.VB_VarHelpID = -1

'***** FUNÇÕES DE INICIALIZAÇÃO DA TELA - INÍCIO *****
Public Sub Form_Load()

Dim lErro As Long
    
On Error GoTo Erro_Form_Load
    
    'Inicializa o Browser
    Set objEventoCliente = New AdmEvento
    Set objEventoBotaoContrato = New AdmEvento
   
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179225)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objContratoPropaganda As ClassContratoPropag) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há uma previsão selecionada exibir seus dados
    If Not (objContratoPropaganda Is Nothing) Then

        'Lê o Contrato passado
        lErro = CF("ContratoPropaganda_Le", objContratoPropaganda)
        If lErro <> SUCESSO And lErro <> 128083 Then gError 128054

        'Contrato não foi cadastrado
        If lErro = 128083 Then gError 128055
        
        'Traz o Contrato para a Tela
        lErro = Traz_ContratoPropaganda_Tela(objContratoPropaganda)
        If lErro <> SUCESSO Then gError 128056

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 128054, 128056
        
        Case 128055
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATOPROPAGANDA_NAO_CADASTRADO", gErr, objContratoPropaganda.lCliente, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179226)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

'*** EVENTOS CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'grava o conteúdo da tela no bd
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 128057

    Call Limpa_ContratoPropaganda_Tela

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 128057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179227)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgResp As VbMsgBoxResult
Dim objContratoPropaganda As New ClassContratoPropag

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se os campos obigatórios estão preenchidos
    If Len(Trim(PeriodoDe.ClipText)) = 0 Then gError 128058
    If Len(Trim(PeriodoAte.ClipText)) = 0 Then gError 128059
    If Len(Trim(Cliente.Text)) = 0 Then gError 128060

    'Preenche o Obj
    lErro = Move_Tela_Memoria(objContratoPropaganda)
    If lErro <> SUCESSO Then gError 128100
    
    'Pede a confirmação da exclusão
    vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CONTRATOPROPAGANDA", objContratoPropaganda.lCliente, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
    
    'se a resposta for não
    If vbMsgResp = vbNo Then Exit Sub

    'transforma o mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Exclui do BD o registro
    lErro = CF("ContratoPropaganda_Exclui", objContratoPropaganda)
    If lErro <> SUCESSO And lErro <> 128096 Then gError 128061

    'se o ContratoPropaganda já foi excluído --> ERRO
    If lErro = 128096 Then gError 128062
    
    'Limpa a tela
    Call Limpa_ContratoPropaganda_Tela
    
    'transforma o mouse em seta padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    'transforma o mouse em seta padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
                
        Case 128058
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODODE_CONTRATOPROPAGANDA_NAO_PREECHIDO", gErr)
                
        Case 128059
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODOATE_CONTRATOPROPAGANDA_NAO_PREENCHIDO", gErr)
        
        Case 128060
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
                
        Case 128061, 128100
        
        Case 128062
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATOPROPAGANDA_NAO_CADASTRADO", gErr, objContratoPropaganda.lCliente, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179228)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa alterações
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 128063

    'limpa a tela toda
    Call Limpa_ContratoPropaganda_Tela
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 128063
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179229)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Call Unload(Me)

End Sub

Private Sub UpDownPeriodoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe_DownClick

    'Se a data está preenchida
    If Len(Trim(PeriodoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoDe, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 128064

    End If

    Exit Sub

Erro_UpDownPeriodoDe_DownClick:

    Select Case gErr

        Case 128064

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179230)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe_UpClick

    'Se a data está preenchida
    If Len(Trim(PeriodoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoDe, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 128065

    End If

    Exit Sub

Erro_UpDownPeriodoDe_UpClick:

    Select Case gErr

        Case 128065

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179231)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_DownClick

    'Se a data está preenchida
    If Len(Trim(PeriodoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoAte, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 128066

    End If

    Exit Sub

Erro_UpDownPeriodoAte_DownClick:

    Select Case gErr

        Case 128066

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179232)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_UpClick

    'Se a data está preenchida
    If Len(Trim(PeriodoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoAte, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 128067

    End If

    Exit Sub

Erro_UpDownPeriodoAte_UpClick:

    Select Case gErr

        Case 128067

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179233)

    End Select

    Exit Sub

End Sub

Private Sub BotaoContrato_Click()
'Realiza a chamada do Browse de ContratoPropaganda

Dim lErro As Long
Dim objContratoPropaganda As New ClassContratoPropag
Dim colSelecao As Collection

On Error GoTo Erro_BotaoContrato_Click

    Call Chama_Tela("ContratoPropagandaLista", colSelecao, objContratoPropaganda, objEventoBotaoContrato)

    Exit Sub
    
Erro_BotaoContrato_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179234)

    End Select

    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    If Len(Trim(Cliente.Text)) <> 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(Cliente.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)
    
    Exit Sub

End Sub
'*** EVENTOS CLICK DOS CONTROLES - FIM ***


'*** EVENTOS CHANGE DOS CONTROLES - INÍCIO ***
Private Sub PeriodoDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PeriodoAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Percentual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTOS CHANGE DOS CONTROLES - FIM ***

'*** EVENTOS VALIDATE DOS CONTROLES - INÍCIO ***
Private Sub PeriodoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoDe_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoDe.Text)
    If lErro <> SUCESSO Then gError 128068

    Exit Sub

Erro_PeriodoDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 128068
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179235)
            
    End Select
    
    Exit Sub

End Sub

Private Sub PeriodoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoAte_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoAte.Text)
    If lErro <> SUCESSO Then gError 128069

    Exit Sub

Erro_PeriodoAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 128069
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179236)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Cliente_Validate

    'se está Preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(Cliente, objCliente, 0)
        If lErro <> SUCESSO Then gError 128070

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 128070

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179237)

    End Select

End Sub

Public Sub Percentual_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Percentual_Validate

    'Verifica a validação da porcentagem
    If Len(Trim(Percentual.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(Percentual)
        If lErro <> SUCESSO Then gError 128071
    End If

    Exit Sub

Erro_Percentual_Validate:

    Cancel = True

    Select Case gErr

        Case 128071

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179238)

    End Select

    Exit Sub

End Sub
'*** EVENTOS VALIDATE DOS CONTROLES - FIM ***

'*** FUNÇÕES DE APOIO A TELA - INÍCIO ***
Public Function Gravar_Registro() As Long
'Realiza a garavação do ContratoPropaganda

Dim lErro As Long
Dim objContratoPropaganda As New ClassContratoPropag

On Error GoTo Erro_Gravar_Registro

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(PeriodoDe.ClipText)) = 0 Then gError 128072
    If Len(Trim(PeriodoAte.ClipText)) = 0 Then gError 128073
    If Len(Trim(Cliente.Text)) = 0 Then gError 128074
    
    'data inicial não pode ser maior que a data final
    If Len(Trim(PeriodoDe.ClipText)) <> 0 And Len(Trim(PeriodoAte.ClipText)) <> 0 Then

         If StrParaDate(PeriodoDe.Text) > StrParaDate(PeriodoAte.Text) Then gError 128101

    End If
    
    'Move da tela para a memória as informações
    lErro = Move_Tela_Memoria(objContratoPropaganda)
    If lErro <> SUCESSO Then gError 128075
    
    'Faz a Chamada da função que irá realizar a gravação do ContratoPropaganda
    lErro = CF("ContratoPropaganda_Grava", objContratoPropaganda)
    If lErro <> SUCESSO Then gError 128076
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 128072
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODODE_CONTRATOPROPAGANDA_NAO_PREENCHIDO", gErr)
        
        Case 128073
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODOATE_CONTRATOPROPAGANDA_NAO_PREENCHIDO", gErr)
        
        Case 128074
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 128075, 128076
        
        Case 128101
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179239)
            
    End Select
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria(objContratoPropaganda As ClassContratoPropag) As Long
'Move da Tela para a Memória

On Error GoTo Erro_Move_Tela_Memoria

    objContratoPropaganda.dtPeriodoDe = MaskedParaDate(PeriodoDe)
    objContratoPropaganda.dtPeriodoAte = MaskedParaDate(PeriodoAte)
    objContratoPropaganda.lCliente = Codigo_Extrai(Cliente.Text)
    objContratoPropaganda.dPercentual = StrParaDbl(Percentual.ClipText) / 100
    objContratoPropaganda.iFilialEmpresa = giFilialEmpresa
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179240)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Limpa_ContratoPropaganda_Tela()

    Call Limpa_Tela(Me)
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
End Sub

Private Function Traz_ContratoPropaganda_Tela(objContratoPropaganda As ClassContratoPropag) As Long
'Traz para a Tela os campos preenchidos com as informações do BD

Dim lErro As Long

On Error GoTo Erro_Traz_ContratoPropaganda_Tela

    'Preenche o Cliente
    Cliente.Text = objContratoPropaganda.lCliente
    Call Cliente_Validate(bSGECancelDummy)
    
    'Preenche o Período
    Call DateParaMasked(PeriodoDe, objContratoPropaganda.dtPeriodoDe)
    Call DateParaMasked(PeriodoAte, objContratoPropaganda.dtPeriodoAte)
    
    'Preenche a Porcentagem
    Percentual.Text = Format(objContratoPropaganda.dPercentual * 100, "Fixed")

    iAlterado = 0
    
    Traz_ContratoPropaganda_Tela = SUCESSO
    
    Exit Function

Erro_Traz_ContratoPropaganda_Tela:

    Traz_ContratoPropaganda_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179241)
            
    End Select
    
    Exit Function
    
End Function
'*** FUNÇÕES DE APOIO A TELA - FIM ***

'*** FUNÇÕES DO SISTEMA DE SETA ***
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objContratoPropaganda As New ClassContratoPropag

On Error GoTo Erro_Tela_Preenche

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    objContratoPropaganda.dtPeriodoDe = CDate(colCampoValor.Item("PeriodoDe").vValor)
    objContratoPropaganda.dtPeriodoAte = CDate(colCampoValor.Item("PeriodoAte").vValor)
    objContratoPropaganda.lCliente = CStr(colCampoValor.Item("Cliente").vValor)
    objContratoPropaganda.dPercentual = CStr(colCampoValor.Item("Percentual").vValor)
    
    'Traz os campos para a Tela
    lErro = Traz_ContratoPropaganda_Tela(objContratoPropaganda)
    If lErro <> SUCESSO Then gError 128077

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 128077

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179242)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
Dim lErro As Long
Dim objContratoPropaganda As New ClassContratoPropag

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada a tela
    sTabela = "ContratoPropaganda"

    'extrai os campos da tela para a memória
    lErro = Move_Tela_Memoria(objContratoPropaganda)
    If lErro <> SUCESSO Then gError 128078

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Cliente", objContratoPropaganda.lCliente, 0, "Cliente"
    colCampoValor.Add "PeriodoDe", objContratoPropaganda.dtPeriodoDe, 0, "PeriodoDe"
    colCampoValor.Add "PeriodoAte", objContratoPropaganda.dtPeriodoAte, 0, "PeriodoAte"
    colCampoValor.Add "Percentual", objContratoPropaganda.dPercentual, 0, "Percentual"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 128078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 179243)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub
'*** FUNÇÕES DO SISTEMA DE SETA - FIM ***

'*** FUNÇÕES DO BROWSER - INÍCIO ***
Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    Cliente.Text = CStr(objCliente.lCodigo)
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoBotaoContrato_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objContratoPropaganda As New ClassContratoPropag

On Error GoTo Erro_objEventoBotaoContrato_evSelecao

    Set objContratoPropaganda = obj1
    
    lErro = Traz_ContratoPropaganda_Tela(objContratoPropaganda)
    If lErro <> SUCESSO Then gError 128079
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoContrato_evSelecao:

    Select Case gErr
        
        Case 128079
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179244)

    End Select

    Exit Sub
    
End Sub
'*** FUNÇÕES DO BROWSER - FIM ***

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoBotaoContrato = Nothing
    Set objEventoCliente = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Contratos de Propaganda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ContratoPropaganda"
    
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
'**** fim do trecho a ser copiado *****

