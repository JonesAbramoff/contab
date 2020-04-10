VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpChequesRecebidos 
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   KeyPreview      =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   6690
   Begin VB.Frame FrameDetalhes 
      Caption         =   "Detalhes"
      Height          =   735
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   4215
      Begin VB.OptionButton DetalhesEspecificados 
         Caption         =   "Especific."
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
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton DetalhesNaoEspecificados 
         Caption         =   "Não Especific."
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
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton DetalhesTodos 
         Caption         =   "Todos"
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
         Left            =   3120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame FrameCliente 
      Caption         =   "Cliente"
      Height          =   735
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   4215
      Begin MSMask.MaskEdBox ClienteDe 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteAte 
         Height          =   315
         Left            =   2685
         TabIndex        =   4
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
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
         Left            =   2265
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   345
         Width           =   360
      End
      Begin VB.Label LabelClienteDe 
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame FrameDataDeposito 
      Caption         =   "Data para Depósito"
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   840
      Width           =   4215
      Begin MSComCtl2.UpDown UpDownDataDepositoDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDepositoDe 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataDepositoAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDepositoAte 
         Height          =   315
         Left            =   2685
         TabIndex        =   2
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataDepositoDe 
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
         Left            =   300
         TabIndex        =   20
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelDataDepositoAte 
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
         Left            =   2265
         TabIndex        =   19
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame FrameLocalizacao 
      Caption         =   "Localização"
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   4215
      Begin VB.OptionButton LocalCaixa 
         Caption         =   "Caixa"
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
         Left            =   1440
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton LocalTodos 
         Caption         =   "Todos"
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
         Left            =   3000
         TabIndex        =   28
         Top             =   720
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton LocalBanco 
         Caption         =   "Banco"
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
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton LocalLoja 
         Caption         =   "Loja"
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
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton LocalBkOffice 
         Caption         =   "Back-Office"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpChequesRecebidos.ctx":0000
      Left            =   1080
      List            =   "RelOpChequesRecebidos.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2670
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpChequesRecebidos.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpChequesRecebidos.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpChequesRecebidos.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpChequesRecebidos.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   11
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
      Left            =   4733
      Picture         =   "RelOpChequesRecebidos.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   945
      Width           =   1605
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
      Left            =   360
      TabIndex        =   14
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpChequesRecebidos"
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

'Obj utilizado para o browser de Clientes
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim giClienteInicial As Integer

Const LOCALIZACAO_CHEQUE_BACKOFFICE = 0
Const LOCALIZACAO_CHEQUE_LOJA = 1
Const LOCALIZACAO_CHEQUE_BANCO = 2
Const LOCALIZACAO_CHEQUE_CAIXA = 3
Const LOCALIZACAO_CHEQUE_TODOS = 4


Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoCliente = New AdmEvento
     
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167526)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Limpa Objetos da memoria
    Set objEventoCliente = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 116500
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116501

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 116501
        
        Case 116500
            '???Luiz: call rotina_erro
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            '???Luiz: call rotina_erro
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167527)

    End Select

    Exit Function

End Function

'???Luiz: não pode existir código entre o trecho comum a ser copiado para todas as telas...
Private Sub BotaoExcluir_Click()
'Aciona a Rotina de exclusão das opções de relatório

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116506

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116507

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click
                 
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116506
            '???Luiz: call rotina_erro
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116507

        Case Else
            '???Luiz: call rotina_erro
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167528)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()
'Aciona rotinas que que checam as opções do relatório e ativam impressão do mesmo

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
    
    'aciona rotina que checa opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116512

    'Chama rotina que excuta a impressão do relatório
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr
        
        Case 116512
        
        Case Else
            '???Luiz: call rotina_erro
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167529)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

 Unload Me
 
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 116502

    'Chama rotina que checa as opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116503

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Aciona rotina que grava opções do relatório
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116504

    'Testa se nome no combo esta igual ao nome em gobjRelOpçoes.sNome
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116505
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116502
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116503 To 116505

        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167530)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116510

    'Coloca os checkbox da tela nas suas opções Default
    LocalTodos.Value = True
    DetalhesTodos.Value = True
    
    'limpa a ComboOpcoes
    ComboOpcoes.Text = ""
    
    'coloca o foco na Combo Opçoes
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 116510
        
        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167531)

    End Select

    Exit Sub

End Sub

Private Sub ClienteAte_LostFocus()
    Call ClienteAte_Validate(bSGECancelDummy)
End Sub
'
'Private Sub ClienteAte_Change()
'
'Dim lErro As Long
'Static sNomeReduzidoParte As String
'
'On Error GoTo Erro_ClienteAte_Change
'
'    'rotina para trazer o nome do cliente com uma parte dos caracteres digitados
'    lErro = CF("Cliente_Pesquisa_NomeReduzido", ClienteAte, sNomeReduzidoParte)
'    If lErro <> SUCESSO Then gError 116542 '??? Luiz: mudar numeração de erro, por favor...
'
'    Exit Sub
'
'Erro_ClienteAte_Change:
'
'    Select Case gErr
'
'        Case 116542
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167532)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo ERRO_ClienteDe_Validate

    If Len(Trim(ClienteAte.ClipText)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteAte, objCliente, 0)
        If lErro <> SUCESSO Then gError 116541
             
    End If
    
    giClienteInicial = 0
    
    Exit Sub

ERRO_ClienteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116541
            '???Luiz: esse erro não deve ser tratado, pois qualquer erro retornado já foi tratada dentro da própria função
            
        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167533)

    End Select

End Sub

Private Sub ClienteDe_LostFocus()
    Call ClienteDe_Validate(bSGECancelDummy)
End Sub

''??? Luiz: eu inclui o código abaixo... favor copiá-lo para ClienteAte_Change também...
'Private Sub ClienteDe_Change()
'
'Dim lErro As Long
'Static sNomeReduzidoParte As String
'
'On Error GoTo Erro_ClienteDe_Change
'
'    'rotina para trazer o nome do cliente com uma parte dos caracteres digitados
'    lErro = CF("Cliente_Pesquisa_NomeReduzido", ClienteDe, sNomeReduzidoParte)
'    If lErro <> SUCESSO Then gError 116511 '??? Luiz: mudar numeração de erro, por favor...
'
'    Exit Sub
'
'Erro_ClienteDe_Change:
'
'    Select Case gErr
'
'        Case 116511
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167534)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub ClienteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo ERRO_ClienteDe_Validate

    If Len(Trim(ClienteDe.ClipText)) > 0 Then
   
        'tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteDe, objCliente, 0)
        If lErro <> SUCESSO Then gError 116540

    End If
    
    giClienteInicial = 1
    
    Exit Sub

ERRO_ClienteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116540
          
        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167535)

    End Select

End Sub

Private Sub ComboOpcoes_Click()
     
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
        
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

     Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataDepositoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDepositoAte)

End Sub

Private Sub DataDepositoAte_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataDepositoAte_Validate

    'Verifica se DataDepositoAte
    If Len(DataDepositoAte.ClipText) > 0 Then
      
        'Verifica Validade da DataDepositoAte
        lErro = Data_Critica(DataDepositoAte.Text)
        If lErro <> SUCESSO Then gError 116546

    End If

    Exit Sub

Erro_DataDepositoAte_Validate:
    
    Cancel = True

    Select Case gErr

        Case 116546
        
        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167536)

    End Select

    Exit Sub

End Sub

Private Sub DataDepositoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDepositoDe)
    
End Sub

Private Sub DataDepositoDe_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataDepositoDe_Validate

    'Verifica se DataDepositoDe foi preenchida
    If Len(DataDepositoDe.ClipText) > 0 Then
                 
        'Verifica Validade da DataDepositoDe
        lErro = Data_Critica(DataDepositoDe.Text)
        If lErro <> SUCESSO Then gError 116547

    End If

    Exit Sub

Erro_DataDepositoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116547

        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167537)

    End Select

    Exit Sub
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_Click()
'???Luiz: incluir tratamento de erro
Dim objCliente As New ClassCliente
Dim colSelecao As Collection

On Error GoTo ErroLabelClienteAte_Click
    giClienteInicial = 0
    
    'Verifica se campo ClienteAte foi preenchido
    If Len(Trim(ClienteAte.ClipText)) > 0 Then
        
        'Preenche objCliente com o cliente que esta na tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteAte.Text)
    
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

    Exit Sub
    
ErroLabelClienteAte_Click:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167538)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)

End Sub

Private Sub LabelClienteDe_Click()
'???Luiz: incluir tratamento de erro
Dim objCliente As New ClassCliente
Dim colSelecao As Collection

On Error GoTo ErroLabelClienteDe_Click

    giClienteInicial = 1
    
    'Verifica se ClienteDe foi preenchido
    If Len(Trim(ClienteDe.ClipText)) > 0 Then
        
        'Preenche o objCliente com o cliente que esta na tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteDe.Text)
    
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

    Exit Sub
    
ErroLabelClienteDe_Click:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167539)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)

End Sub

Private Sub UpDownDataDepositoAte_DownClick()
'Diminui DataDepositoAte em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataDepositoAte_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataDepositoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116528

    Exit Sub

Erro_UpDownDataDepositoAte_DownClick:

    Select Case gErr

        Case 116528
            DataDepositoAte.SetFocus

        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167540)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataDepositoAte_UpClick()
'Aumenta DataDepositoAte em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataDepositoAte_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataDepositoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116530

    Exit Sub

Erro_UpDownDataDepositoAte_UpClick:

    Select Case gErr

        Case 116530
            DataDepositoAte.SetFocus

        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167541)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataDepositoDe_DownClick()
'Diminui DataDepositoDe em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataDepositoDe_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataDepositoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116527

    Exit Sub

Erro_UpDownDataDepositoDe_DownClick:

    Select Case gErr

        Case 116527
            DataDepositoDe.SetFocus

        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167542)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDepositoDe_UpClick()
'Aumenta DataDepositoDe em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataDepositoDe_UpClick

    'Aciona rotna que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataDepositoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116529

    Exit Sub

Erro_UpDownDataDepositoDe_UpClick:

    Select Case gErr

        Case 116529
            DataDepositoDe.SetFocus

        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167543)

    End Select

    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Verifica se a tecla F3 (Browser) foi acionada, e qual Browser ela deve trazer
    If KeyCode = KEYCODE_BROWSER Then
        
        'Verifica se o campo atual é o ClienteDe ou o ClienteAté
        If Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        End If
    
    End If

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sDetalhes As String
Dim sLocalizacao As String
Dim sDataInic As String
Dim sDataFim As String

On Error GoTo Erro_PreencherRelOp

    'Verifica Parametros , e formata os mesmos
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sDataInic, sDataFim)
    If lErro <> SUCESSO Then gError 116513
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116514
   
    'Inclui parametro de ClienteDe
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 116515
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116532
    
    'Inclui parametro de ClienteAte
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 116516
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteAte.ClipText)
    If lErro <> AD_BOOL_TRUE Then gError 116533
    
    '???Luiz: pq você não formata a data dentro da função Formata_E_Critica_Parametros?
    'Inclui parametro de DataDepositoDe
    lErro = objRelOpcoes.IncluirParametro("DINI", sDataInic)
    If lErro <> AD_BOOL_TRUE Then gError 116517

    'Inclui parametro de DataDepositoAte
    lErro = objRelOpcoes.IncluirParametro("DFIM", sDataFim)
    If lErro <> AD_BOOL_TRUE Then gError 116518
   
    'Verifica opção de Localização que esta marcada
    If LocalLoja.Value = True Then
        sLocalizacao = LOCALIZACAO_CHEQUE_LOJA

    ElseIf LocalBanco.Value = True Then
        sLocalizacao = LOCALIZACAO_CHEQUE_BANCO
        
    ElseIf LocalBkOffice.Value = True Then
        sLocalizacao = LOCALIZACAO_CHEQUE_BACKOFFICE
        
    ElseIf LocalCaixa.Value = True Then
        sLocalizacao = LOCALIZACAO_CHEQUE_CAIXA
        
    ElseIf LocalTodos.Value = True Then
        sLocalizacao = LOCALIZACAO_CHEQUE_TODOS
        
    End If
    
    'Inclui parametro de Localizacao
    lErro = objRelOpcoes.IncluirParametro("NLOCALIZACAO", sLocalizacao)
    If lErro <> AD_BOOL_TRUE Then gError 116534

    'Verifica qual opção de DETALHES esta marcada
    If DetalhesEspecificados.Value = True Then
        sDetalhes = CHEQUE_ESPECIFICADO '???Luiz: use a constante já existente CHEQUE_ESPECIFICADO
    
    ElseIf DetalhesNaoEspecificados.Value = True Then
        sDetalhes = CHEQUE_NAO_ESPECIFICADO '???Luiz: use a constante já existente CHEQUE_NAO_ESPECIFICADO
        
    ElseIf DetalhesTodos.Value = True Then
        sDetalhes = "2" '???Luiz: use a constante já existente CHEQUE_NAO_ESPECIFICADO
        
    End If
    
    'Inclui parametro de Detalhes
    lErro = objRelOpcoes.IncluirParametro("NDETALHES", sDetalhes)
    If lErro <> AD_BOOL_TRUE Then gError 116535 '???Luiz: erro não tratado
    
    '???Luiz: esse filtro só deve ser passado se o a filial ativa for diferente de EMPRESA_TODA
    'Inclui Parametro Filial Empresa
    If giFilialEmpresa <> EMPRESA_TODA Then
        lErro = objRelOpcoes.IncluirParametro("NFILIALEMPRESA", CStr(giFilialEmpresa))
        If lErro <> AD_BOOL_TRUE Then gError 116536 '???Luiz: erro não tratado
    
    End If
    
    'Aciona Rotina que monta_expressão que será usada para gerar relatório
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_I, sCliente_F, sLocalizacao, sDetalhes, sDataInic, sDataFim)
    If lErro <> SUCESSO Then gError 116519
        
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 116513 To 116519
        
        Case 116532 To 116536
                       
        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167544)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    'inicializa bSGECancelDummy como Falsa
    bSGECancelDummy = False
       
    'Limpa a Tela
    lErro = Limpa_Tela
    If lErro <> SUCESSO Then gError 116792
              
    'Carrega parametros do relatorio gravado
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 116520
            
    'pega parâmetro Pedido Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro Then gError 116521
    
    'Preenche campo ClienteDe
    ClienteDe.Text = sParam
    
    'verifica validade do ClienteDe
    Call ClienteDe_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then gError 116683
    
    'pega parâmetro Pedido Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro Then gError 116524
    
    'Preenche campo ClienteAte
    ClienteAte.Text = sParam
    
    'verifica validade do ClienteAte
    Call ClienteAte_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then gError 116684
                
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError 116522

    'Preenche campo DataDepositoDe
    Call DateParaMasked(DataDepositoDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 116523

    'Preenche campo DataDepositoAte
    Call DateParaMasked(DataDepositoAte, CDate(sParam))
    
    'Pega Localizacao e Exibe
    lErro = objRelOpcoes.ObterParametro("NLOCALIZACAO", sParam)
    If lErro <> SUCESSO Then gError 116537 '???Luiz: erro não tratado

    'Verifica qual a opção de Localizção do relatório carregado
    Select Case sParam
        
        Case LOCALIZACAO_CHEQUE_LOJA '???Luiz: substituir por constantes
            LocalLoja.Value = True
            
        Case LOCALIZACAO_CHEQUE_BANCO
            LocalBanco.Value = True
            
        Case LOCALIZACAO_CHEQUE_BACKOFFICE
            LocalBkOffice.Value = True
        
        Case Else
            LocalTodos = True
    
    End Select
   
    'Pega especificação e a Exibe
    lErro = objRelOpcoes.ObterParametro("NDETALHES", sParam)
    If lErro <> SUCESSO Then gError 116539 '???Luiz: erro não tratado
    
    'Verifica qual a opção de DETALHES do relatório carregado
    Select Case sParam
        
        Case CHEQUE_NAO_ESPECIFICADO '???Luiz: substituir por constantes
            DetalhesNaoEspecificados.Value = True
            
            
        Case CHEQUE_ESPECIFICADO
            DetalhesEspecificados.Value = True
            
        Case Else
            DetalhesTodos.Value = True
    
    End Select
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 116520 To 116524
        
        Case 116683
            ClienteDe.Text = ""
                                   
        Case 116684
            ClienteAte.Text = ""
            
        Case 116792
                                   
        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167545)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_I As String, sCliente_F As String, sDataInic As String, sDataFim As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'Formata ClienteDe
    If ClienteDe.ClipText <> "" Then
        sCliente_I = CStr(LCodigo_Extrai(ClienteDe.Text))
    Else
        sCliente_I = ""
    End If
    
    'Formata ClienteAte
    If ClienteAte.ClipText <> "" Then
        sCliente_F = CStr(LCodigo_Extrai(ClienteAte.Text))
    Else
        sCliente_F = ""
    End If
    
    'verifica se ClienteDe é maior que o ClienteAte
    If Trim(ClienteDe.ClipText) <> "" And Trim(ClienteAte.ClipText) <> "" Then
    
         If CLng(sCliente_I) > CLng(sCliente_F) Then gError 116525
         
    End If
    
    'verifica se DataDepositoDe é maior que  DataDepositoAte
    If Trim(DataDepositoDe.ClipText) <> "" Then
    
        sDataInic = DataDepositoDe.Text
        If Trim(DataDepositoAte.ClipText) <> "" Then
    
            sDataFim = DataDepositoAte.Text
            'se DataDepositoDe > DataDepositoAte => ERRO
            If CDate(sDataInic) > CDate(sDataFim) Then gError 116526
    
        Else
        
            'preenche sDataFim com DATA NULA
            sDataFim = CStr(DATA_NULA)
            
        End If
        
    Else
    
        'preenche sDataInic com DATA NULA
        sDataInic = CStr(DATA_NULA)
        If Trim(DataDepositoAte.ClipText) <> "" Then
    
            sDataFim = DataDepositoAte.Text
                
        Else
        
            'preenche sDataFim com DATA NULA
            sDataFim = CStr(DATA_NULA)
            
        End If
        
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
            
        Case 116525
            '???Luiz:call
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEINICIAL_MAIOR_CLIENTEFINAL", gErr)
            ClienteDe.SetFocus
        
        Case 116526
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDepositoDe.SetFocus
               
         Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167546)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_I As String, sCliente_F As String, sLocalizacao As String, sDetalhes As String, sDataInic As String, sDataFim As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    'Verifica se campo ClienteDe foi preenchido
    If Trim(ClienteDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressao o Valor de ClienteDe
        sExpressao = sExpressao & "Cliente >= " & sCliente_I
        
    End If

    'Verifica se campo ClienteAte foi preenchido
    If Trim(ClienteAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de ClienteAte
        sExpressao = sExpressao & "Cliente <= " & sCliente_F

    End If
    
    'Verifica se campo DataDepositoDe foi preenchido
    If Trim(DataDepositoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de DataDepositoDe
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(sDataInic))

    End If
    
    'Verifica se campo DataDepositoAte foi preenchido
    If Trim(DataDepositoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de DataDepositoAte
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(sDataFim))

    End If
    
    'Verifica se Localização selecionada é diferente do Localização TODOS
    If LocalTodos.Value = False Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o Valor do Localização
        sExpressao = sExpressao & "Localizacao = " & sLocalizacao
    
    End If
    
    'Verifica se Detalhes selecionado é diferente de Detalhes TODOS
    If DetalhesTodos.Value = False Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o Valor de DETALHES '???Luiz: ao invés de NDETALHES, passar Detalhes
        sExpressao = sExpressao & "Detalhes = " & sDetalhes
    
    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de Filial Empresa '???Luiz: ao invés de filialempresa>=, deveria ser filialempresa=
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
    
    End If
    
    'Verifica se a expressão foi preenchido com algum filtro
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            '???Luiz: call
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167547)

    End Select

    Exit Function

End Function

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

On Error GoTo ErroobjEventoCliente_evSelecao
'???Luiz: incluir tratamento de erro

    Set objCliente = obj1
    
    'se controle atual é o ClienteDe
    If giClienteInicial = 1 Then
        
        'Preenche campo ClienteDe
        ClienteDe.Text = CStr(objCliente.lCodigo)
                
        'verifica validade de ClienteDe
        Call ClienteDe_Validate(bSGECancelDummy)
    
    'Se controle atual é o ClienteAte
    Else
       
       'Preenche campo ClienteAte
       ClienteAte.Text = CStr(objCliente.lCodigo)
              
       'Verifica Validade de ClienteAte
       Call ClienteAte_Validate(bSGECancelDummy)
    
    End If

    Me.Show

    Exit Sub
    
ErroobjEventoCliente_evSelecao:
        
    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167548)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela()
'Limpa os campos da tela , quando é chamada uma opção de relatorio para a tela

On Error GoTo Erro_Limpa_Tela

    'Limpa campos de data
    DataDepositoDe.Text = "  /  /  "
    DataDepositoAte.Text = "  /  /  "
    
    'Limpa campos de Cliente
    ClienteDe.Text = ""
    ClienteAte.Text = ""
    
    'Seta valores default para Options Buttons
    LocalTodos.Value = True
    DetalhesTodos.Value = True
    
    Exit Function

Erro_Limpa_Tela:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167549)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Relação de Cheques Recebidos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpChequesRecebidos"
    
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

'???Luiz: esse código "genérico" deve ficar junto com o restante
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub
'???Luiz: esse código "genérico" deve ficar junto com o restante
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

'???Luiz: esse código "genérico" deve ficar junto com o restante
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

'???Luiz: esse código "genérico" deve ficar junto com o restante
Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

'???Luiz: esse código "genérico" deve ficar junto com o restante
Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub


