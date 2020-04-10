VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRelCuponsFiscais 
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   KeyPreview      =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   6690
   Begin VB.CheckBox ExibirCancelados 
      Caption         =   "Exibir Cupons Cancelados"
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
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelCuponsFiscais.ctx":0000
      Left            =   1080
      List            =   "RelOpRelCuponsFiscais.ctx":0002
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRelCuponsFiscais.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpRelCuponsFiscais.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpRelCuponsFiscais.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpRelCuponsFiscais.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   15
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
      Left            =   4740
      Picture         =   "RelOpRelCuponsFiscais.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   945
      Width           =   1605
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   240
      TabIndex        =   25
      Top             =   840
      Width           =   4215
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   690
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
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
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
      Begin VB.Label LabelDataDe 
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
         TabIndex        =   29
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelDataAte 
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
         TabIndex        =   28
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   1680
      Width           =   4215
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   4
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelCaixaDe 
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
         Height          =   195
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelCaixaAte 
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
         Height          =   195
         Left            =   2280
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame FrameCupomFiscal 
      Caption         =   "Cupom Fiscal"
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   4215
      Begin MSMask.MaskEdBox CupomFiscalDe 
         Height          =   315
         Left            =   810
         TabIndex        =   5
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CupomFiscalAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   6
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCupomFiscalAte 
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
         Height          =   195
         Left            =   2280
         TabIndex        =   21
         Top             =   345
         Width           =   360
      End
      Begin VB.Label LabelCupomFiscalDe 
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
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame FrameVendedor 
      Caption         =   "Vendedor"
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   4215
      Begin MSMask.MaskEdBox VendedorDe 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   8
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedorDe 
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
         Height          =   195
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelVendedorAte 
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
         Height          =   195
         Left            =   2280
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.CheckBox ExibirItens 
      Caption         =   "Exibir Itens"
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
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
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
      TabIndex        =   31
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelCuponsFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giVendedorInicial As Integer
Dim giCaixaInicial As Integer

'Obj utilizado para o browser de Caixas
Private WithEvents objEventoCaixa As AdmEvento
Attribute objEventoCaixa.VB_VarHelpID = -1

'Obj utilizado para o browser de Vendedores
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoCaixa = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172506)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Limpa Objetos da memoria
    Set objEventoCaixa = Nothing
    Set objEventoVendedor = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 116576
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116577

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 116577
        
        Case 116576
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172507)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Aciona a Rotina de exclusão das opções de relatório

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116578

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116579

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click
            
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116578
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116579

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172508)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()
'Aciona rotinas que que checam as opções do relatório e ativam impressão do mesmo

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
    
    'aciona rotina que checa opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116582

    'Verifica qual tsk deve ser executado
    If ExibirItens.Value = vbChecked Then
        gobjRelatorio.sNomeTsk = "RECPFIT"
    Else
        gobjRelatorio.sNomeTsk = "RECPF"
    End If
    
    'Chama rotina que excuta a impressão do relatório
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr
        
        Case 116582
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172509)

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
    If ComboOpcoes.Text = "" Then gError 116583

    'Chama rotina que checa as opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116584

    'Seta o nome da opção que será gravado como o nome que esta na comboOpções
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Aciona rotina que grava opções do relatório
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116585

    'Testa se nome no combo esta igual ao nome em gobjRelOpçoes.sNome
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116586
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116583
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116584 To 116586

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172510)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Chama função que limpa Relatório
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116587
    
    'Desmarca ExibirCancelados e ExibirItens
    ExibirCancelados.Value = Unchecked
        
    ExibirItens.Value = Unchecked
    
    'Limpa a combo opções
    ComboOpcoes.Text = ""
    
    'Seta o foco na ComboOpções
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 116587
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172511)

    End Select

    Exit Sub

End Sub

Private Sub CaixaAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(CaixaAte)
End Sub

Private Sub CaixaAte_Validate(Cancel As Boolean)
'Verifica validade de CaixaAte

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaAte_Validate

    giCaixaInicial = 0

    If Len(Trim(CaixaAte.Text)) > 0 Then

        'instancia o obj
        Set objCaixa = New ClassCaixa

        'preenche o obj c/ o cod e filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler a Caixa (Código ou nome)
        lErro = CF("TP_Caixa_Le1", CaixaAte, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116628

        'código inexistente
        If lErro = 116175 Then gError 116629

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116629

    End If
 
    Exit Sub

Erro_CaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116628

        Case 116629
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, CaixaAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172512)

    End Select

End Sub

Private Sub CaixaDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(CaixaDe)
End Sub

Private Sub CaixaDe_Validate(Cancel As Boolean)
'Verifica validade de CaixaDe

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaDe_Validate
        
    giCaixaInicial = 1

    If Len(Trim(CaixaDe.Text)) > 0 Then
        
        'instancia o obj
        Set objCaixa = New ClassCaixa
        
        'preenche o obj c/ o cod e filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler Caixa (Código ou nome)
        lErro = CF("TP_Caixa_Le1", CaixaDe, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116630

        'código inexistente
        If lErro = 116175 Then gError 116631

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116631

    End If
    
    Exit Sub

Erro_CaixaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116630

        Case 116631
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, CaixaDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172513)

    End Select

End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Private Sub CupomFiscalAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(CupomFiscalAte)
End Sub

Private Sub CupomFiscalAte_Validate(Cancel As Boolean)
'Verifica validade de CupomFiscalDe

Dim lErro As Long

On Error GoTo Erro_CupomFiscalAte_Validate

    If Len(Trim(CupomFiscalAte.ClipText)) > 0 Then

        'Verifica se valor digitado é valido
        lErro = Long_Critica(CupomFiscalAte.Text)
        If lErro <> SUCESSO Then gError 102470

    End If

    Exit Sub

Erro_CupomFiscalAte_Validate:

    Cancel = True

    Select Case gErr

        Case 102470
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172514)

    End Select
    
    Exit Sub

End Sub

Private Sub CupomFiscalDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(CupomFiscalDe)
End Sub

Private Sub CupomFiscalDe_Validate(Cancel As Boolean)
'Verifica validade de CupomFiscalDe

Dim lErro As Long

On Error GoTo Erro_CupomFiscalDe_Validate

    If Len(Trim(CupomFiscalDe.ClipText)) > 0 Then

        'Verifica se valor digitado é valido
        lErro = Long_Critica(CupomFiscalDe.Text)
        If lErro <> SUCESSO Then gError 102471

    End If

    Exit Sub

Erro_CupomFiscalDe_Validate:

    Cancel = True

    Select Case gErr

        Case 102471
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172515)

    End Select

    Exit Sub
    
End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte)
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se DataAte
    If Len(DataAte.ClipText) > 0 Then

        'Verifica Validade da DataAte
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 116589

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116589
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172516)

    End Select

    Exit Sub
    
End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe)
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se DataDe foi preenchida
    If Len(DataDe.ClipText) > 0 Then

        'Verifica Validade da DataDe
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 116590

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116590

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172517)

    End Select

    Exit Sub
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelCaixaAte_Click()

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaAte_Click

    giCaixaInicial = 0
    
    If Len(Trim(CaixaAte.Text)) > 0 Then
        'Preenche com a caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
    End If
    
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa)
    
    Else
    
        'Chama Tela Caixa
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)
    
    End If
    
    Exit Sub

Erro_LabelCaixaAte_Click:
   
   Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172518)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelCaixaDe_Click()

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaDe_Click

    giCaixaInicial = 1
    
    If Len(Trim(CaixaDe.Text)) > 0 Then
        'Preenche com a caixa  da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
    End If
    
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa)
    
    Else
    
        'Chama Tela de caixa
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)
    
    End If
    
    Exit Sub

Erro_LabelCaixaDe_Click:
   
   Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172519)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelCupomFiscalAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelCupomFiscalAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCupomFiscalDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelCupomFiscalDe, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorAte_Click()
'Aciona Browser de Vendedores
Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 0
    
    If Len(Trim(VendedorAte.ClipText)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorAte.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub LabelVendedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelVendedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorDe_Click()
'Aciona o Browser de Vendedores
Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 1
    
    If Len(Trim(VendedorDe.ClipText)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorDe.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub LabelVendedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelVendedorDe, Button, Shift, X, Y)
End Sub

Private Sub objEventoCaixa_evSelecao(obj1 As Object)

Dim objCaixa As ClassCaixa

On Error GoTo Erro_objEventoCaixa_evSelecao

    Set objCaixa = obj1

    'se controle atual é o CaixaDe
    If giCaixaInicial = 1 Then

        'Preenche campo CaixaDe
        CaixaDe.Text = CStr(objCaixa.iCodigo)
      
        Call CaixaDe_Validate(bSGECancelDummy)

    'Se controle atual é o CaixaAte
    Else

       'Preenche campo CaixaAte
       CaixaAte.Text = CStr(objCaixa.iCodigo)
      
       Call CaixaAte_Validate(bSGECancelDummy)

    End If

    Me.Show

    Exit Sub

Erro_objEventoCaixa_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172520)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)
'Preenche campo vendedor com valor trazido pelo Browser

Dim objVendedor As ClassVendedor

On Error GoTo Erro_objEventoVendedor_evSelecao
    
    Set objVendedor = obj1
    
    'verifica qual campo deve ser preenchido VendedorDe ou VendedorAte
    If giVendedorInicial = 1 Then
        
        'Preenche o campo VendedorDe
        VendedorDe.Text = CStr(objVendedor.iCodigo)
              
        'verifica validade do campo VendedorDe
        Call VendedorDe_Validate(bSGECancelDummy)
    
    Else
        
        'Preenche o campo VendedorAte
        VendedorAte.Text = CStr(objVendedor.iCodigo)
              
        'verifica validade do campo VendedorAte
        Call VendedorAte_Validate(bSGECancelDummy)
    
    End If

    Me.Show

    Exit Sub
    
Erro_objEventoVendedor_evSelecao:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172521)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataAte_DownClick()
'Diminui DataAte em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116591

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 116591
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172522)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataAte_UpClick()
'Aumenta DataAte em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116594

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 116594
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172523)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataDe_DownClick()
'Diminui DataDe em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116592

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 116592
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172524)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataDe_UpClick()
'Aumenta DataDe em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116593

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 116593
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172525)

    End Select

    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Verifica se a tecla F3 (Browser) foi acionada, e qual Browser ela deve trazer
    If KeyCode = KEYCODE_BROWSER Then
        
        'Verifica qual browser deve ser acionado
        If Me.ActiveControl Is CaixaDe Then
            Call LabelCaixaDe_Click
        ElseIf Me.ActiveControl Is CaixaAte Then
            Call LabelCaixaAte_Click
        ElseIf Me.ActiveControl Is VendedorDe Then
            Call LabelVendedorDe_Click
        ElseIf Me.ActiveControl Is VendedorAte Then
            Call LabelVendedorAte_Click
        End If
    
    End If

End Sub

Private Sub VendedorAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(VendedorAte)
End Sub

Private Sub VendedorAte_Validate(Cancel As Boolean)
'verifica validade do campo VendedorAte

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorAte_Validate

    If Len(Trim(VendedorAte.ClipText)) > 0 Then
   
        'Tenta ler o vendedor (Código)
        lErro = TP_Vendedor_Le2(VendedorAte, objVendedor, 0)
        If lErro <> SUCESSO And lErro <> 25011 And lErro <> 25013 Then gError 116626
        If lErro <> SUCESSO Then gError 116627

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 116626

        Case 116627
             Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, VendedorAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172526)

    End Select
    
    Exit Sub
    
End Sub

Private Sub VendedorDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(VendedorDe)
End Sub

Private Sub VendedorDe_Validate(Cancel As Boolean)
'verifica validade do campo VendedorDe

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorDe_Validate

    If Len(Trim(VendedorDe.ClipText)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorDe, objVendedor, 0)
        If lErro <> SUCESSO And lErro <> 25011 And lErro <> 25013 Then gError 116624
        If lErro <> SUCESSO Then gError 116625

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116624
        
        Case 116625
             Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, VendedorDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172527)

    End Select
    
End Sub

Private Function Formata_E_Critica_Parametros(sCaixa_I As String, sCaixa_F As String, sVendedor_I As String, sVendedor_F As String, sDataInic As String, sDataFim As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    '****** CAIXA *******************
    'Verifica preenchimento de CaixaDe
    If Len(Trim(CaixaDe.ClipText)) > 0 Then
        sCaixa_I = CStr(Codigo_Extrai(CaixaDe.Text))
    Else
        sCaixa_I = ""
    End If
    
    'verifica preenchimento de CaixaAte
    If Len(Trim(CaixaAte.ClipText)) > 0 Then
        sCaixa_F = CStr(Codigo_Extrai(CaixaAte.Text))
    Else
        sCaixa_F = ""
    End If
        
    If sCaixa_I <> "" And sCaixa_F <> "" Then
    
        'se CaixaDe > CaixaAte => Erro
        If StrParaInt(sCaixa_I) > StrParaInt(sCaixa_F) Then gError 116595
    
    End If
    '************************************
        
    '********** DATA *************
    'formata datas e verifica se DataDe é maior que  DataAte
    If Trim(DataDe.ClipText) <> "" Then
        sDataInic = DataDe.Text
    Else
        sDataInic = DATA_NULA
    End If
    
    If Trim(DataAte.ClipText) <> "" Then
        sDataFim = DataAte.Text
    Else
        sDataFim = DATA_NULA
    End If
    
    If sDataInic <> DATA_NULA And sDataFim <> DATA_NULA Then
        'se dataDe > DataAte => Erro
        If CDate(sDataInic) > CDate(sDataFim) Then gError 116596
    End If
    '******************************
    
    'verifica se o CupomFiscalDe é maior que o CupomFiscalAte
    If Trim(CupomFiscalDe.ClipText) <> "" And Trim(CupomFiscalAte.ClipText) <> "" Then
    
         If StrParaInt(CupomFiscalDe.Text) > StrParaInt(CupomFiscalAte.Text) Then gError 116597
    
    End If
        
    '********** VENDEDOR ******************
    'verifica preenchimento de VendedorDe
    If Len(Trim(VendedorDe.Text)) > 0 Then
        sVendedor_I = CStr(Codigo_Extrai(VendedorDe.Text))
    Else
        sVendedor_I = ""
    End If
    
    'verifica preenchimento de VendedorAte
    If Len(Trim(VendedorAte.Text)) > 0 Then
        sVendedor_F = CStr(Codigo_Extrai(VendedorAte.Text))
    Else
        sVendedor_F = ""
    End If
    
    If sVendedor_I <> "" And sVendedor_F <> "" Then
    
        'verifica se o VendedorDe é maior que o VendedorAte
        If StrParaInt(sVendedor_I) > CInt(sVendedor_F) Then gError 116598
        
    End If
    '********************************************
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
            
        Case 116595
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXAINICIAL_MAIOR_CAIXAFINAL", gErr)
            CaixaDe.SetFocus
        
        Case 116596
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
               
         Case 116597
            Call Rotina_Erro(vbOKOnly, "ERRO_CUPOMFISCAL_INICIAL_MAIOR", gErr)
            CupomFiscalDe.SetFocus
            
         Case 116598
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR2", gErr)
            VendedorDe.SetFocus
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172528)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCaixa_I As String
Dim sCaixa_F As String
Dim sVendedor_I As String
Dim sVendedor_F As String
Dim sDataInic As String
Dim sDataFim As String

On Error GoTo Erro_PreencherRelOp

    'Verifica Parametros , e formata os mesmos
    lErro = Formata_E_Critica_Parametros(sCaixa_I, sCaixa_F, sVendedor_I, sVendedor_F, sDataInic, sDataFim)
    If lErro <> SUCESSO Then gError 116599
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116600
   
    'Inclui parametro de CaixaDe
    lErro = objRelOpcoes.IncluirParametro("NCAIXAINIC", sCaixa_I)
    If lErro <> AD_BOOL_TRUE Then gError 116601
    
    lErro = objRelOpcoes.IncluirParametro("TCAIXAINIC", Trim(CaixaDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116677
    
    'Inclui parametro de CaixaAte
    lErro = objRelOpcoes.IncluirParametro("NCAIXAFIM", sCaixa_F)
    If lErro <> AD_BOOL_TRUE Then gError 116602
    
     lErro = objRelOpcoes.IncluirParametro("TCAIXAFIM", Trim(CaixaAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116678
    
    'Inclui parametro de DataDe
    lErro = objRelOpcoes.IncluirParametro("DINI", sDataInic)
    If lErro <> AD_BOOL_TRUE Then gError 116603

    'Inclui parametro de DataAte
    lErro = objRelOpcoes.IncluirParametro("DFIM", sDataFim)
    If lErro <> AD_BOOL_TRUE Then gError 116604
   
    'Inclui parametro de CupomFiscalDe
    lErro = objRelOpcoes.IncluirParametro("NCUPOMFISCALINIC", CupomFiscalDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116605
    
    'Inclui parametro de CupomFiscalAte
    lErro = objRelOpcoes.IncluirParametro("NCUPOMFISCALFIM", CupomFiscalAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116606
    
    'Inclui parametro de VendedorDe
    lErro = objRelOpcoes.IncluirParametro("NVENDEDORINIC", sVendedor_I)
    If lErro <> AD_BOOL_TRUE Then gError 116607
    
     lErro = objRelOpcoes.IncluirParametro("TVENDEDORINIC", Trim(VendedorDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116679
    
    'Inclui parametro de VendedorAte
    lErro = objRelOpcoes.IncluirParametro("NVENDEDORFIM", sVendedor_F)
    If lErro <> AD_BOOL_TRUE Then gError 116608
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDORFIM", Trim(VendedorDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116680
   
    'Inclui parametro de Exibir Cancelados
    lErro = objRelOpcoes.IncluirParametro("NEXIBIRCANCELADOS", CStr(ExibirCancelados.Value))
    If lErro <> AD_BOOL_TRUE Then gError 116609
  
    'Inclui parametro de Exibir Itens
    lErro = objRelOpcoes.IncluirParametro("NEXIBIRITENS", CStr(ExibirItens.Value))
    If lErro <> AD_BOOL_TRUE Then gError 116610
    
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        'Inclui Parametro Filial Empresa
        lErro = objRelOpcoes.IncluirParametro("NFILIALEMPRESA", CStr(giFilialEmpresa))
        If lErro <> AD_BOOL_TRUE Then gError 116611
    
    End If
    
    'Aciona Rotina que monta_expressão que será usada para gerar relatório
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sDataInic, sDataFim, sCaixa_I, sCaixa_F, sVendedor_I, sVendedor_F)
    If lErro <> SUCESSO Then gError 116612
        
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr


    Select Case gErr

        Case 116599 To 116612
        
        Case 116677 To 116680
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172529)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sDataInic As String, sDataFim As String, sCaixa_I As String, sCaixa_F As String, sVendedor_I As String, sVendedor_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    'Verifica se campo CaixaDe foi preenchido
    If Trim(CaixaDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressao o Valor de CaixaDe
        sExpressao = sExpressao & "Caixa >= " & sCaixa_I
        
    End If

    'Verifica se campo CaixaAte foi preenchido
    If Trim(CaixaAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de CaixaAte
        sExpressao = sExpressao & "Caixa <= " & sCaixa_F

    End If
    
    'Verifica se campo DataDe foi preenchido
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de DataDe
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(sDataInic))

    End If
    
    'Verifica se campo DataAte foi preenchido
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de DataAte
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(sDataFim))

    End If
    
    'Verifica se campo CupomFiscalDe foi preenchido
    If Trim(CupomFiscalDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de CupomFiscalDe
        sExpressao = sExpressao & "CupomFiscal >= " & Forprint_ConvInt(CInt(CupomFiscalDe.Text))

    End If
    
    'Verifica se campo CupomFiscalAte foi preenchido
    If Trim(CupomFiscalAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de CupomFiscalate
        sExpressao = sExpressao & "CupomFiscal <= " & Forprint_ConvInt(CInt(CupomFiscalAte.Text))

    End If
    
    'Verifica se campo VendedorDe foi preenchido
    If Trim(VendedorDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de VendedorDe
        sExpressao = sExpressao & "Vendedor >= " & sVendedor_I

    End If
    
    'Verifica se campo VendedorAte foi preenchido
    If Trim(VendedorAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de VendedorAte
        sExpressao = sExpressao & "Vendedor >= " & sVendedor_F

    End If
    
    'Verifica se Exibir Cancelados esta marcado
    If CInt(ExibirCancelados.Value) = 1 Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o Valor do Exibir Cancelados
        sExpressao = sExpressao & "Status <> " & STATUS_CANCELADO
    
    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de Filial Empresa
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172530)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    'Inicializa variavel bSGECancelDummy
    bSGECancelDummy = False

    'Limpa a Tela
    lErro = Limpa_Tela
    If lErro <> SUCESSO Then gError 116793
    
    'Carrega parametros do relatorio gravado
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 116613
            
    'pega parâmetro Caixa Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAINIC", sParam)
    If lErro Then gError 116614
    
    'Preenche campo CaixaDe
    CaixaDe.Text = sParam
        
    'verifica validade de CaixaDe
    Call CaixaDe_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then gError 116689
    
    'pega parâmetro Caixa Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAFIM", sParam)
    If lErro Then gError 116615
    
    'Preenche campo CaixaAte
    CaixaAte.Text = sParam
    
    'verifica validade de CaixaAte
    Call CaixaAte_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then gError 116690
                
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError 116616

    'Preenche campo DataDe
    Call DateParaMasked(DataDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 116617

    'Preenche campo DataAte
    Call DateParaMasked(DataAte, CDate(sParam))
    
    'Pega parametro CupomFiscalDe e o Exibe
    lErro = objRelOpcoes.ObterParametro("NCUPOMFISCALINIC", sParam)
    If lErro <> SUCESSO Then gError 116618

    'Preenche campo CupomFiscalDe
    CupomFiscalDe.PromptInclude = False
    CupomFiscalDe.Text = sParam
    CupomFiscalDe.PromptInclude = True
    
    'Pega parametro CupomFiscalAte e o Exibe
    lErro = objRelOpcoes.ObterParametro("NCUPOMFISCALFIM", sParam)
    If lErro <> SUCESSO Then gError 116619

    'Preenche campo CupomFiscalAte
    CupomFiscalAte.PromptInclude = False
    CupomFiscalAte.Text = sParam
    CupomFiscalAte.PromptInclude = True
    
    'Pega parametro VendedorDe e o Exibe
    lErro = objRelOpcoes.ObterParametro("NVENDEDORINIC", sParam)
    If lErro <> SUCESSO Then gError 116620

    'Preenche campo VendedorDe
    VendedorDe.Text = sParam
    
    'verifica validade de VendedorDe
    Call VendedorDe_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then gError 116691
        
    'Pega parametro VendedorAte e o Exibe
    lErro = objRelOpcoes.ObterParametro("NVENDEDORFIM", sParam)
    If lErro <> SUCESSO Then gError 116621

    'Preenche campo VendedorAte
    VendedorAte.Text = sParam
        
    'verifica validade de VendedorAte
    Call VendedorAte_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then gError 116692
    
    'Pega Exibir Cancelados e exibe
    lErro = objRelOpcoes.ObterParametro("NEXIBIRCANCELADOS", sParam)
    If lErro <> SUCESSO Then gError 116622
    
    'verifica se Exibir Cancelados esta marcado no relatorio carregado
    If sParam = MARCADO Then
        
        ExibirCancelados.Value = Checked
    
    Else
    
        ExibirCancelados.Value = Unchecked
        
    End If
    
    'Pega Exibir Itens e exibe
    lErro = objRelOpcoes.ObterParametro("NEXIBIRITENS", sParam)
    If lErro <> SUCESSO Then gError 116623
            
    'verifica se Exibir Itens esta marcado no relatorio carregado
    If sParam = MARCADO Then
    
        ExibirItens.Value = Checked
           
    Else
    
        ExibirItens.Value = Unchecked
        
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 116613 To 116623
        
        Case 116793
        
        Case 116689
            CaixaDe.Text = ""
            
        Case 116690
            CaixaAte.Text = ""
            
        Case 116691
            VendedorDe.Text = ""
            
        Case 116692
            VendedorAte.Text = ""
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172531)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela()
'Limpa os campos da tela , quando é chamada uma opção de relatorio para a tela

On Error GoTo Erro_Limpa_Tela

    'Limpa campos de data
    DataDe.Text = "  /  /  "
    DataAte.Text = "  /  /  "
    
    'Limpa campos de Caixa
    CaixaDe.Text = ""
    CaixaAte.Text = ""
    
   'Limpa campos de Cupom Fiscal
    CupomFiscalDe.PromptInclude = False
    CupomFiscalDe.Text = ""
    CupomFiscalDe.PromptInclude = True
    
    CupomFiscalAte.PromptInclude = False
    CupomFiscalAte.Text = ""
    CupomFiscalAte.PromptInclude = True
    
    'Limpa campos de Vendedor
    VendedorDe.Text = ""
    VendedorAte.Text = ""
    
    Exit Function

Erro_Limpa_Tela:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172532)

    End Select

    Exit Function

End Function


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Cupons Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelCuponsFiscais"

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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub LabelCaixaAte_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCaixaAte, Source, X, Y)
End Sub

Private Sub LabelCaixaDe_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCaixaDe, Source, X, Y)
End Sub

Private Sub LabelCupomFiscalAte_DragDrop(Source As Control, X As Single, Y As Single)
     Call Controle_DragDrop(LabelCupomFiscalAte, Source, X, Y)
End Sub

Private Sub LabelCupomFiscalDe_DragDrop(Source As Control, X As Single, Y As Single)
 Call Controle_DragDrop(LabelCupomFiscalDe, Source, X, Y)
End Sub

Private Sub LabelVendedorAte_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelVendedorAte, Source, X, Y)
End Sub

Private Sub LabelVendedorDe_DragDrop(Source As Control, X As Single, Y As Single)
 Call Controle_DragDrop(LabelVendedorDe, Source, X, Y)
End Sub
