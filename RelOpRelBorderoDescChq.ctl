VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRelBorderoDescChq 
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   6705
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
      Picture         =   "RelOpRelBorderoDescChq.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   945
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpRelBorderoDescChq.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpRelBorderoDescChq.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpRelBorderoDescChq.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRelBorderoDescChq.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   4215
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   690
         TabIndex        =   13
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   315
         Left            =   2685
         TabIndex        =   15
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
         TabIndex        =   17
         Top             =   345
         Width           =   360
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
         TabIndex        =   16
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelBorderoDescChq.ctx":0A96
      Left            =   1080
      List            =   "RelOpRelBorderoDescChq.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   270
      Width           =   2670
   End
   Begin VB.Frame FrameBordero 
      Caption         =   "Borderô"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1690
      Width           =   4215
      Begin MSMask.MaskEdBox BorderoDe 
         Height          =   315
         Left            =   810
         TabIndex        =   6
         Top             =   285
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox BorderoAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   7
         Top             =   285
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelBorderoDe 
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
         TabIndex        =   9
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelBorderoAte 
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
         TabIndex        =   8
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.CheckBox DetalharCheques 
      Caption         =   "Detalhar Cheques"
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
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Frame FrameContaCorrente 
      Caption         =   "Conta Corrente"
      Height          =   1185
      Left            =   240
      TabIndex        =   0
      Top             =   2540
      Width           =   4215
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   1245
         TabIndex        =   3
         Top             =   675
         Width           =   2550
      End
      Begin VB.OptionButton ContaCorrenteApenas 
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
         TabIndex        =   2
         Top             =   712
         Width           =   1050
      End
      Begin VB.OptionButton ContaCorrenteTodas 
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
         TabIndex        =   1
         Top             =   285
         Width           =   900
      End
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
      TabIndex        =   24
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelBorderoDescChq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giBorderoInicial As Integer

'Obj utilizado para o browser de Borderos
Private WithEvents objEventoBordero As AdmEvento
Attribute objEventoBordero.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa objeto usado pelo Browser
    Set objEventoBordero = New AdmEvento
    
    'Lê as contas correntes  com codigo e o nome reduzido existentes no BD e carrega na ComboBox
    lErro = Carrega_ContaCorrente()
    If lErro <> SUCESSO Then gError 116746
       
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 116746

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172400)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Limpa Objetos da memoria
    Set objEventoBordero = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 116747
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116748

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 116748
        
        Case 116747
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172401)

    End Select

    Exit Function

End Function

Private Sub BorderoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(BorderoAte)
    
End Sub

Private Sub BorderoAte_Validate(Cancel As Boolean)
'verifica validade do campo BorderoAte
Dim lErro As Long

On Error GoTo Erro_BorderoAte_Validate

    If Len(Trim(BorderoAte.ClipText)) > 0 Then
        
        'verifica validade do campo BorderoAte
        lErro = Long_Critica(BorderoAte.Text)
        If lErro <> SUCESSO Then gError 116749
        
    End If

    giBorderoInicial = 1

    Exit Sub

Erro_BorderoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116749

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172402)

    End Select

End Sub

Private Sub BorderoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(BorderoDe)
    
End Sub

Private Sub BorderoDe_Validate(Cancel As Boolean)
'verifica validade do campo BorderoDe
Dim lErro As Long

On Error GoTo Erro_Borderode_Validate

    If Len(Trim(BorderoDe.ClipText)) > 0 Then

        'verifica validade do campo BorderoAte
        lErro = Long_Critica(BorderoDe.Text)
        If lErro <> SUCESSO Then gError 116751
        
    End If

    giBorderoInicial = 1

    Exit Sub

Erro_Borderode_Validate:

    Cancel = True

    Select Case gErr

        Case 116751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172403)

    End Select

End Sub

Private Sub BotaoExcluir_Click()
'Aciona a Rotina de exclusão das opções de relatório

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116753

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116754

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Aciona Rotinas para Limpar Tela
        Call BotaoLimpar_Click
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116753
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116754

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172404)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()
'Aciona rotinas que que checam as opções do relatório e ativam impressão do mesmo

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
    
    'aciona rotina que checa opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116755
    
    'Se não deve detalhar o relatório
    If CInt(DetalharCheques.Value) = vbUnchecked Then
    
        'Guarda o nome do tsk com os dados resumidos
        gobjRelatorio.sNomeTsk = "BDDCCHRS"
        
    'Senão, ou seja, se é para detalhar o relatório
    Else
    
        'Guarda o nome do tsk com os dados detalhados
        gobjRelatorio.sNomeTsk = "BDDCCHDT"

    End If
    
    
    'Chama rotina que excuta a impressão do relatório
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr
        
        Case 116755
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172405)

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
    If ComboOpcoes.Text = "" Then gError 116756

    'Chama rotina que checa as opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116757

    'Seta o nome da opção que será gravado como o nome que esta na comboOpções
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Aciona rotina que grava opções do relatório
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116758

    'Testa se nome no combo esta igual ao nome em gobjRelOpçoes.sNome
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116759
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116756
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116757 To 116759

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172406)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
'Aciona Rotinas de Limpeza de tela
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Chama função que limpa Relatório
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116760
          
    'Limpa o campo ComboOpcoes
    ComboOpcoes.Text = ""
    
    'coloca Conta Corrente Todas como Default Novamente
    ContaCorrenteTodas.Value = True
    
    'Seta o foco na ComboOpções
    ComboOpcoes.SetFocus
    
    'Limpa o Combo de conta corrente
    ContaCorrente.Text = ""
                
    'Desmarca opção Detalhar Cheques
    DetalharCheques.Value = Unchecked
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 116760
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172407)

    End Select

    Exit Sub
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub ContaCorrente_Change()

    ContaCorrenteApenas.Value = True
    
End Sub

Private Sub ContaCorrente_DropDown()
    
    ContaCorrenteApenas.Value = True

End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)
'Verifica validade da conta corrente selecionada

Dim lErro As Long

On Error GoTo Erro_ContaCorrente_Validate

    lErro = CF("ContaCorrente_Bancaria_ValidaCombo", ContaCorrente)
    If lErro <> SUCESSO Then gError 116761

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True

    Select Case gErr

        Case 116761

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172408)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long

On Error GoTo Erro_Carrega_ContaCorrente

    lErro = CF("ContasCorrentes_Bancarias_CarregaCombo", ContaCorrente)
    If lErro <> SUCESSO Then gError 116762
    
    Carrega_ContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_ContaCorrente:

    Carrega_ContaCorrente = gErr

    Select Case gErr

        Case 116762

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172409)

    End Select

    Exit Function

End Function

Private Sub ContaCorrenteTodas_Click()

    'Limpa campo Conta Corrente
    ContaCorrente.Text = ""
    
    'Marca novamente opção Conta Corrente Todas
    ContaCorrenteTodas.Value = True
    
End Sub

Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Verifica validade de DataAte
Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se DataAte foi preenchida
    If Len(DataAte.ClipText) > 0 Then

        sDataFim = DataAte.Text
        
        'Verifica Validade da DataAte
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 116763

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116763
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172410)

    End Select

    Exit Sub
    
End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se DataDe foi preenchida
    If Len(DataDe.ClipText) > 0 Then

        sDataInic = DataDe.Text
        
        'Verifica Validade da DataDe
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 116764

    End If

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True

    Select Case gErr

        Case 116764

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172411)

    End Select

    Exit Sub
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelBorderoAte_Click()
'Aciona o Browser de Borderos
Dim objBorderoDescChq As New ClassBorderoDescChq
Dim colSelecao As Collection

On Error GoTo Erro_LabelBorderoAte_Click

    giBorderoInicial = 0
    
    If Len(Trim(BorderoAte.ClipText)) > 0 Then
        'Preenche com o Bordero da tela
        objBorderoDescChq.lNumBordero = BorderoAte.Text
    End If
    
    'Chama Tela BorderoChequeLojaLista
    Call Chama_Tela("BorderoDescChqLista", colSelecao, objBorderoDescChq, objEventoBordero)

    Exit Sub
    
Erro_LabelBorderoAte_Click:
  Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172412)

    End Select

    Exit Sub
End Sub

Private Sub LabelBorderoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelBorderoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelBorderoDe_Click()
'Aciona o Browser de Borderos
Dim objBorderoDescChq As New ClassBorderoDescChq
Dim colSelecao As Collection

On Error GoTo Erro_LabelBorderoDe_Click

    giBorderoInicial = 1
    
    If Len(Trim(BorderoDe.ClipText)) > 0 Then
        'Preenche com o Bordero da tela
        objBorderoDescChq.lNumBordero = StrParaLong(BorderoDe.Text)
    End If
    
    'Chama Tela BorderoChequeLojaLista
     Call Chama_Tela("BorderoDescChqLista", colSelecao, objBorderoDescChq, objEventoBordero)

    Exit Sub
    
Erro_LabelBorderoDe_Click:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172413)

    End Select

    Exit Sub

End Sub

Private Sub LabelBorderoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelBorderoDe, Button, Shift, X, Y)
End Sub

Private Sub objEventoBordero_evSelecao(obj1 As Object)
'Preenche campo Bordero com valor trazido pelo Browser
Dim objBorderoDescChq As ClassBorderoDescChq

    Set objBorderoDescChq = obj1
    
    'verifica qual campo deve ser preenchido BorderoDe ou BorderoAte
    If giBorderoInicial = 1 Then
        
        'Preenche o campo BOrderoDe
        BorderoDe.PromptInclude = False
        BorderoDe.Text = CStr(objBorderoDescChq.lNumBordero)
        BorderoDe.PromptInclude = True
        
        'verifica validade do campo BorderoDe
        BorderoDe_Validate (bSGECancelDummy)
    
    Else
        
        'Preenche o campo BorderoAte
        BorderoAte.PromptInclude = False
        BorderoAte.Text = CStr(objBorderoDescChq.lNumBordero)
        BorderoAte.PromptInclude = True
        
        'verifica validade do campo BorderoAte
        BorderoAte_Validate (bSGECancelDummy)
    
    End If

    Me.Show

    Exit Sub
    
End Sub

Private Sub UpDownDataAte_DownClick()
'Diminui DataAte em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116765

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 116765
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172414)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataAte_UpClick()
'Aumenta DataAte em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116766

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 116766
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172415)

    End Select

    Exit Sub
End Sub

Private Sub UpDownDataDe_DownClick()
'Diminui DataDe em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116767

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 116767
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172416)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataDe_UpClick()
'Aumenta DataDe em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116768

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 116768
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172417)

    End Select

    Exit Sub
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Verifica se a tecla F3 (Browser) foi acionada
    If KeyCode = KEYCODE_BROWSER Then
        
        'Verifica qual browser deve ser acionado
        If Me.ActiveControl Is BorderoDe Then
            Call LabelBorderoDe_Click
        ElseIf Me.ActiveControl Is BorderoAte Then
            Call LabelBorderoAte_Click
        End If
    
    End If
    
End Sub

Private Function Formata_E_Critica_Parametros(sDataInic As String, sDataFim As String, sContaCorrente As String) As Long
'Formata e verifica validade das opções passadas para gerar o relatório
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata datas e verifica se DataDe é maior que  DataAte
    If Trim(DataDe.ClipText) <> "" Then
    
        sDataInic = DataDe.Text
        'verifica se DataAte foi preenchida
        If Trim(DataAte.ClipText) <> "" Then
    
            sDataFim = DataAte.Text
            'se DataDe for é maior que DataAte => ERRO
            If CDate(sDataInic) > CDate(sDataFim) Then gError 116769
    
        Else
        
            sDataFim = CStr(DATA_NULA)
            
        End If
        
    Else
        'preenche sDataInic com DATA NULA
        sDataInic = CStr(DATA_NULA)
        'verifica se DataAte foi preenchida
        If Trim(DataAte.ClipText) <> "" Then
    
            sDataFim = DataAte.Text
                
        Else
            'Preenche DataAte com DATA NULA
            sDataFim = CStr(DATA_NULA)
            
        End If
        
    End If
             
    'verifica se o BorderoDe é maior que o BorderoAte
    If Trim(BorderoDe.ClipText) <> "" And Trim(BorderoAte.ClipText) <> "" Then
    
         If CLng(BorderoDe.Text) > CLng(BorderoAte.Text) Then gError 116770
    
    End If
    
    'Verifica qual opção de Conta Corrente esta marcada
    If ContaCorrenteTodas.Value <> True Then
           
        'verifica se foi preenchida combo conta corrente
        If ContaCorrente.Text = "" Then gError 116778
                
        'Formata Conta Corrente
        sContaCorrente = Codigo_Extrai(ContaCorrente.Text)
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
               
        Case 116769
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
               
        Case 116770
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERO_INICIAL_MAIOR", gErr)
            BorderoDe.SetFocus
         
        Case 116778
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PREENCHIDA2", gErr)
            ContaCorrente.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172418)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sDestino As String
Dim sDataInic As String
Dim sDataFim As String
Dim sContaCorrente As String

On Error GoTo Erro_PreencherRelOp

    'Verifica Parametros , e formata os mesmos
    lErro = Formata_E_Critica_Parametros(sDataInic, sDataFim, sContaCorrente)
    If lErro <> SUCESSO Then gError 116771
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116772
   
    'Inclui parametro de DataDe
    lErro = objRelOpcoes.IncluirParametro("DINI", sDataInic)
    If lErro <> AD_BOOL_TRUE Then gError 116773

    'Inclui parametro de DataAte
    lErro = objRelOpcoes.IncluirParametro("DFIM", sDataFim)
    If lErro <> AD_BOOL_TRUE Then gError 116774
       
    'Inclui parametro de BorderoDe
    lErro = objRelOpcoes.IncluirParametro("NBORDEROINIC", BorderoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116775
    
    'Inclui parametro de BorderoAte
    lErro = objRelOpcoes.IncluirParametro("NBORDEROFIM", BorderoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116776
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        'Inclui Parametro Filial Empresa
        lErro = objRelOpcoes.IncluirParametro("NFILIALEMPRESA", CStr(giFilialEmpresa))
        If lErro <> AD_BOOL_TRUE Then gError 116777
    
    End If
            
    'Inclui parametro de Conta Corrente
    lErro = objRelOpcoes.IncluirParametro("NCONTACORRENTE", sContaCorrente)
    If lErro <> AD_BOOL_TRUE Then gError 116779
    
    'Inclui parametro de Conta Corrente
    lErro = objRelOpcoes.IncluirParametro("TCONTACORRENTE", ContaCorrente.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116779
    
    'Inclui parametro de Detalhar Cheques
    lErro = objRelOpcoes.IncluirParametro("NDETALHARCHEQUES", CInt(DetalharCheques.Value))
    If lErro <> AD_BOOL_TRUE Then gError 116780
        
    'Aciona Rotina que monta_expressão que será usada para gerar relatório
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sDataInic, sDataFim, sContaCorrente)
    If lErro <> SUCESSO Then gError 116781
        
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 116771 To 116777
        
        Case 116779 To 116781
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172419)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sDataInic As String, sDataFim As String, sContaCorrente As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

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
        
    'Verifica se campo BorderoDe foi preenchido
    If Trim(BorderoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de BorderoDe
        sExpressao = sExpressao & "Bordero >= " & Forprint_ConvLong(CLng(BorderoDe.Text))

    End If
    
    'Verifica se campo BorderoAte foi preenchido
    If Trim(BorderoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de BorderoAte
        sExpressao = sExpressao & "Bordero <= " & Forprint_ConvLong(CLng(BorderoAte.Text))

    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de Filial Empresa
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
    
    End If
    
    'Verifica se a opção de Conta Corrente selecionada é diferente de TODAS
    If ContaCorrenteTodas.Value = False Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o Valor de Conta Corrente
        sExpressao = sExpressao & "ContaCorrente = " & sContaCorrente
    
    End If
    
'    'Verifica se a Checkbox Detalhar Cheques esta selecionada
'    If DetalharCheques.Value = Checked Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        'Inclui na expressão o Valor de Conta Corrente
'        sExpressao = sExpressao & "DetalharCheques = " & Forprint_ConvInt(CInt(DetalharCheques.Value))
'
'    End If
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172420)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    'Limpa a Tela
    lErro = Limpa_Tela
    If lErro <> SUCESSO Then gError 116786
    
    'Carrega parametros do relatorio gravado
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 116782
            
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError 116783

    'Preenche campo DataDe
    Call DateParaMasked(DataDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 116784

    'Preenche campo DataAte
    Call DateParaMasked(DataAte, CDate(sParam))
        
    'Pega parametro BorderoDe e o Exibe
    lErro = objRelOpcoes.ObterParametro("NBORDEROINIC", sParam)
    If lErro <> SUCESSO Then gError 116785

    'Preenche campo BorderoDe
    BorderoDe.PromptInclude = False
    BorderoDe.Text = sParam
    BorderoDe.PromptInclude = True
        
    'Pega parametro BorderoAte e o Exibe
    lErro = objRelOpcoes.ObterParametro("NBORDEROFIM", sParam)
    If lErro <> SUCESSO Then gError 116788

    'Preenche campo BorderoAte
    BorderoAte.PromptInclude = False
    BorderoAte.Text = sParam
    BorderoAte.PromptInclude = True
                
    'Pega a Conta Corrente e Exibe
    lErro = objRelOpcoes.ObterParametro("NCONTACORRENTE", sParam)
    If lErro <> SUCESSO Then gError 116789
    
    'Verifica se conta corrente foi selecionada
    If Len(Trim(sParam)) > 0 Then
    
        ContaCorrente.Text = sParam
        Call ContaCorrente_Validate(bSGECancelDummy)
        If bSGECancelDummy = True Then gError 116790
        
    End If
    
    'Pega Detalhar Cheques e exibe
    lErro = objRelOpcoes.ObterParametro("NDETALHARCHEQUES", sParam)
    If lErro <> SUCESSO Then gError 116791
    
    'verifica se Detalhar Cheques esta marcado no relatorio carregado
    If sParam <> "" Then
        
        DetalharCheques.Value = Checked
    
    Else
    
        DetalharCheques.Value = Unchecked
        
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 116782 To 116785
        
        Case 116788 To 116791
        
        Case 116786
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172421)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela()
'Limpa os campos da tela , quando é chamada uma opção de relatorio para a tela

On Error GoTo Erro_Limpa_Tela

    'Limpa campos de data
    DataDe.Text = "  /  /  "
    DataAte.Text = "  /  /  "
    
    'Limpa campos de Bordero
    BorderoDe.PromptInclude = False
    BorderoDe.Text = ""
    BorderoDe.PromptInclude = True
    
    BorderoAte.PromptInclude = False
    BorderoAte.Text = ""
    BorderoAte.PromptInclude = True
    
    'Seta valores default para Options Buttons e checkBox
    ContaCorrenteTodas.Value = True
    DetalharCheques = Unchecked

    Exit Function
    
Erro_Limpa_Tela:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172422)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Bordero Desconto Cheques"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelBorderoDescChq"

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

Private Sub LabelBorderoAte_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelBorderoAte, Source, X, Y)
End Sub
Private Sub LabelBorderoDe_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelBorderoDe, Source, X, Y)
End Sub

