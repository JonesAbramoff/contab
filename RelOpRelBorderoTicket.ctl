VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRelBorderoTicket 
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   6705
   Begin VB.Frame FrameAdministradora 
      Caption         =   "Administradora"
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   4215
      Begin MSMask.MaskEdBox AdministradoraDe 
         Height          =   315
         Left            =   810
         TabIndex        =   5
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AdministradoraAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   6
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelAdministradoraAte 
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
         TabIndex        =   24
         Top             =   345
         Width           =   360
      End
      Begin VB.Label LabelAdministradoraDe 
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
         TabIndex        =   23
         Top             =   345
         Width           =   315
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
      Picture         =   "RelOpRelBorderoTicket.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpRelBorderoTicket.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpRelBorderoTicket.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpRelBorderoTicket.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRelBorderoTicket.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   4215
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   16
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
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   17
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelBorderoTicket.ctx":0A96
      Left            =   1080
      List            =   "RelOpRelBorderoTicket.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2670
   End
   Begin VB.Frame FrameBordero 
      Caption         =   "Borderô"
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   4215
      Begin MSMask.MaskEdBox BorderoDe 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox BorderoAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   4
         Top             =   285
         Width           =   1215
         _ExtentX        =   2143
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   345
         Width           =   360
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
      TabIndex        =   21
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelBorderoTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variaveis utilizadas para os Browsers identificarem qual campo deve ser preenchido pelos mesmos
Dim giBorderoInicial As Integer
Dim giAdministradoraInicial As Integer

'Obj utilizado para o browser de Borderos
Private WithEvents objEventoBordero As AdmEvento
Attribute objEventoBordero.VB_VarHelpID = -1

'Obj utilizado para o browser de Administradoras
Private WithEvents objEventoAdministradora As AdmEvento
Attribute objEventoAdministradora.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializa objetos usados pelos Browsers
    Set objEventoBordero = New AdmEvento
    Set objEventoAdministradora = New AdmEvento
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172423)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Limpa Objetos da memoria
    Set objEventoBordero = Nothing
    Set objEventoAdministradora = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 116633
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116634

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 116634
        
        Case 116633
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172424)

    End Select

    Exit Function

End Function

Private Sub AdministradoraAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(AdministradoraAte)

End Sub

Private Sub AdministradoraAte_Validate(Cancel As Boolean)
''verifica validade do campo AdministradoraAte
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_AdministradoraAte_Validate

    If Len(Trim(AdministradoraAte.ClipText)) > 0 Then

        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
        objAdmMeioPagto.iCodigo = AdministradoraAte.Text
        'Tenta ler o Administradora (Código ou Nome Reduzido)
        lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO And lErro <> 116678 And lErro <> 116680 Then gError 116658
        If lErro <> SUCESSO Then gError 116659

    End If

    giAdministradoraInicial = 1

    Exit Sub

Erro_AdministradoraAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116658

        Case 116659
             Call Rotina_Erro(vbOKOnly, "ERRO_ADMINISTRADORA_NAO_CADASTRADA", gErr, AdministradoraAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172425)

    End Select

End Sub

Private Sub AdministradoraDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(AdministradoraDe)

End Sub

Private Sub AdministradoraDe_Validate(Cancel As Boolean)
''verifica validade do campo AdministradoraDe
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_Administradorade_Validate

    If Len(Trim(AdministradoraDe.ClipText)) > 0 Then

        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
        objAdmMeioPagto.iCodigo = AdministradoraDe.Text
        'Tenta ler o Administradora
        lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO And lErro <> 116678 And lErro <> 116680 Then gError 116652
        If lErro <> SUCESSO Then gError 116653

    End If

    giAdministradoraInicial = 1

    Exit Sub

Erro_Administradorade_Validate:

    Cancel = True

    Select Case gErr

        Case 116652

        Case 116653
             Call Rotina_Erro(vbOKOnly, "ERRO_ADMNISTRADORA_NAO_CADASTRADA", gErr, AdministradoraDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172426)

    End Select
    
End Sub

Private Sub BorderoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(BorderoAte)
    
End Sub

Private Sub BorderoAte_Validate(Cancel As Boolean)
'verifica validade do campo BorderoAte
Dim lErro As Long

On Error GoTo Erro_BorderoAte_Validate

    If Len(Trim(BorderoAte.ClipText)) > 0 Then
        
        'verifica validade de BorderoAte
        Long_Critica (BorderoAte.Text)
        If lErro <> SUCESSO Then gError 116656
        
    End If

    giBorderoInicial = 1

    Exit Sub

Erro_BorderoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116656

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172427)

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

        'verifica validade de BorderoDe
        Long_Critica (BorderoDe.Text)
        If lErro <> SUCESSO Then gError 116654
        
    End If

    giBorderoInicial = 1

    Exit Sub

Erro_Borderode_Validate:

    Cancel = True

    Select Case gErr

        Case 116654

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172428)

    End Select

End Sub

Private Sub BotaoExcluir_Click()
'Aciona a Rotina de exclusão das opções de relatório

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116635

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116636

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 116637
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116635
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116636, 116637

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172429)

    End Select

    Exit Sub
End Sub

Private Sub BotaoExecutar_Click()
'Aciona rotinas que que checam as opções do relatório e ativam impressão do mesmo

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
    
    'aciona rotina que checa opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116639

    'Chama rotina que excuta a impressão do relatório
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr
        
        Case 116639
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172430)

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
    If ComboOpcoes.Text = "" Then gError 116640

    'Chama rotina que checa as opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116641

    'Seta o nome da opção que será gravado como o nome que esta na comboOpções
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Aciona rotina que grava opções do relatório
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116642

    'Testa se nome no combo esta igual ao nome em gobjRelOpçoes.sNome
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116643
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116640
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116641 To 116643

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172431)

    End Select

    Exit Sub
End Sub

Private Sub BotaoLimpar_Click()
'Aciona Rotinas de Limpeza de tela
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Chama função que limpa Relatório
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116644
          
    'Limpa o campo ComboOpcoes
    ComboOpcoes.Text = ""
    
    'Seta o foco na ComboOpções
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 116644
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172432)

    End Select

    Exit Sub
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

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
        If lErro <> SUCESSO Then gError 116647

    End If

    Exit Sub

Erro_DataAte_Validate:

    DataAte.SelStart = 0
    DataAte.SelLength = Len(DataAte)
    Cancel = True

    Select Case gErr

        Case 116647
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172433)

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
        If lErro <> SUCESSO Then gError 116646

    End If

    Exit Sub

Erro_DataDe_Validate:

    DataDe.SelStart = 0
    DataDe.SelLength = Len(DataDe)
    Cancel = True


    Select Case gErr

        Case 116646

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172434)

    End Select

    Exit Sub
            
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelAdministradoraAte_Click()
'Aciona o Browser de Administradoras
Dim objAdministradora As New ClassAdmMeioPagto
Dim colSelecao As Collection

On Error GoTo Erro_LabelAdministradoraAte_Click

    giAdministradoraInicial = 0

    If Len(Trim(AdministradoraAte.ClipText)) > 0 Then
        'Preenche com o Administradora da tela
        objAdministradora.iCodigo = LCodigo_Extrai(AdministradoraAte.Text)
    End If

    'Chama Tela AdministradoraLista
    Call Chama_Tela("AdmMeioPagtoLista", colSelecao, objAdministradora, objEventoAdministradora)
    
    Exit Sub
    
Erro_LabelAdministradoraAte_Click:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172435)

    End Select

    Exit Sub

End Sub

Private Sub LabelAdministradoraAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelAdministradoraAte, Button, Shift, X, Y)
End Sub

Private Sub LabelAdministradoraDe_Click()
'Aciona o Browser de Administradoras
Dim objAdministradora As New ClassAdmMeioPagto
Dim colSelecao As Collection

On Error GoTo Erro_LabelAdministradoraDe_Click
    
    giAdministradoraInicial = 1

    If Len(Trim(AdministradoraDe.ClipText)) > 0 Then
        'Preenche com o Administradora da tela
        objAdministradora.iCodigo = LCodigo_Extrai(AdministradoraDe.Text)
    End If

    'Chama Tela AdministradoraLista
    Call Chama_Tela("AdmMeioPagtoLista", colSelecao, objAdministradora, objEventoAdministradora)

    Exit Sub
    
Erro_LabelAdministradoraDe_Click:
  
  Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172436)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelAdministradoraDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelAdministradoraDe, Button, Shift, X, Y)
End Sub

Private Sub LabelBorderoAte_Click()
'Aciona o Browser de Borderos
Dim objBorderoTicket As New ClassBorderoValeTicket
Dim colSelecao As Collection

On Error GoTo Erro_LabelBorderoAte_Click
    
    giBorderoInicial = 0
    
    If Len(Trim(BorderoAte.ClipText)) > 0 Then
        'Preenche com o Bordero da tela
        objBorderoTicket.lNumBordero = BorderoAte.Text
    End If
    
    'Chama Tela BorderoLista
    Call Chama_Tela("BorderoValeTicketLista", colSelecao, objBorderoTicket, objEventoBordero)
    
    Exit Sub
      
Erro_LabelBorderoAte_Click:

  Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172437)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelBorderoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelBorderoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelBorderoDe_Click()
'Aciona o Browser de Borderos
Dim objBorderoTicket As New ClassBorderoValeTicket
Dim colSelecao As Collection

On Error GoTo Erro_LabelBorderoDe_Click

    giBorderoInicial = 1
    
    If Len(Trim(BorderoDe.ClipText)) > 0 Then
        'Preenche com o Bordero da tela
        objBorderoTicket.lNumBordero = BorderoDe.Text
    End If
    
    'Chama Tela BorderoLista
    Call Chama_Tela("BorderoValeTicketLista", colSelecao, objBorderoTicket, objEventoBordero)
  
    Exit Sub
    
Erro_LabelBorderoDe_Click:

  Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172438)

    End Select

    Exit Sub

End Sub

Private Sub LabelBorderoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelBorderoDe, Button, Shift, X, Y)
End Sub

Private Sub objEventoAdministradora_evSelecao(obj1 As Object)
Dim objAdministradora As ClassAdmMeioPagto

    Set objAdministradora = obj1

    'se controle atual é o AdministradoraDe
    If giAdministradoraInicial = 1 Then

        'Preenche campo AdministradoraDe
        AdministradoraDe.Text = CStr(objAdministradora.iCodigo)


        Call AdministradoraDe_Validate(bSGECancelDummy)

    'Se controle atual é o AdministradoraAte
    Else

       'Preenche campo AdministradoraAte
       AdministradoraAte.Text = CStr(objAdministradora.iCodigo)

       Call AdministradoraAte_Validate(bSGECancelDummy)

    End If

    Me.Show

    Exit Sub
    
End Sub

Private Sub objEventoBordero_evSelecao(obj1 As Object)
'Preenche campo Bordero com valor trazido pelo Browser
Dim objBorderoTicket As ClassBorderoValeTicket

    Set objBorderoTicket = obj1
    
    'verifica qual campo deve ser preenchido BorderoDe ou BorderoAte
    If giBorderoInicial = 1 Then
        
        'Preenche o campo BOrderoDe
        BorderoDe.PromptInclude = False
        BorderoDe.Text = CStr(objBorderoTicket.lNumBordero)
        BorderoDe.PromptInclude = True
        
        'verifica validade do campo BorderoDe
        BorderoDe_Validate (bSGECancelDummy)
    
    Else
        
        'Preenche o campo BorderoAte
        BorderoAte.PromptInclude = False
        BorderoAte.Text = CStr(objBorderoTicket.lNumBordero)
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
    If lErro <> SUCESSO Then gError 116650

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 116650
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172439)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataAte_UpClick()
'Aumenta DataAte em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116651

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 116651
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172440)

    End Select

    Exit Sub
End Sub

Private Sub UpDownDataDe_DownClick()
'Diminui DataDe em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116648

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 116648
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172441)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataDe_UpClick()
'Aumenta DataDe em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116649

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 116649
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172442)

    End Select

    Exit Sub
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Verifica se a tecla F3 (Browser) foi acionada, e qual Browser ela deve trazer
    If KeyCode = KEYCODE_BROWSER Then
        
        'Verifica qual browser deve ser acionado
        If Me.ActiveControl Is BorderoDe Then
            Call LabelBorderoDe_Click
        ElseIf Me.ActiveControl Is BorderoAte Then
            Call LabelBorderoAte_Click
        ElseIf Me.ActiveControl Is AdministradoraDe Then
            Call LabelAdministradoraDe_Click
        ElseIf Me.ActiveControl Is AdministradoraAte Then
            Call LabelAdministradoraAte_Click
        End If
    
    End If
    
End Sub

Private Function Formata_E_Critica_Parametros(sAdmMeioPagto_I As String, sAdmMeioPagto_F As String, sDataInic As String, sDataFim As String) As Long
'Formata e verifica validade das opções passadas para gerar o relatório
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'Formata AdministradoraDe
    If AdministradoraDe.ClipText <> "" Then
        sAdmMeioPagto_I = CStr(LCodigo_Extrai(AdministradoraDe.Text))
    Else
        sAdmMeioPagto_I = ""
    End If

    'Formata AdministradoraAte
    If AdministradoraAte.ClipText <> "" Then
        sAdmMeioPagto_F = CStr(LCodigo_Extrai(AdministradoraAte.Text))
    Else
        sAdmMeioPagto_F = ""
    End If

    'verifica se AdministradoraDe é maior que o AdministradoraAte
    If Trim(AdministradoraDe.ClipText) <> "" And Trim(AdministradoraAte.ClipText) <> "" Then

         If CLng(sAdmMeioPagto_I) > CLng(sAdmMeioPagto_F) Then gError 116660

    End If
    
    'formata datas e verifica se DataDe é maior que  DataAte
    If Trim(DataDe.ClipText) <> "" Then
    
        sDataInic = DataDe.Text
        'verifica se DataAte foi preenchida
        If Trim(DataAte.ClipText) <> "" Then
    
            sDataFim = DataAte.Text
            'se DataDe for é maior que DataAte => ERRO
            If CDate(sDataInic) > CDate(sDataFim) Then gError 116661
    
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
    
         If CLng(BorderoDe.Text) > CLng(BorderoAte.Text) Then gError 116662
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
            
        Case 116660
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMINICIAL_MAIOR_ADMFINAL", gErr)
            AdministradoraDe.SetFocus
        
        Case 116661
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
               
        Case 116662
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERO_INICIAL_MAIOR", gErr)
            BorderoDe.SetFocus
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172443)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sExibirCancelados As String
Dim sExibirItens As String
Dim sAdmMeioPagto_I As String
Dim sAdmMeioPagto_F As String
Dim sDataInic As String
Dim sDataFim As String

On Error GoTo Erro_PreencherRelOp

    'Verifica Parametros , e formata os mesmos
    lErro = Formata_E_Critica_Parametros(sAdmMeioPagto_I, sAdmMeioPagto_F, sDataInic, sDataFim)
    If lErro <> SUCESSO Then gError 116663
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116664
   
    'Inclui parametro de AdministradoraDe
    lErro = objRelOpcoes.IncluirParametro("NADMINISTRADORAINIC", sAdmMeioPagto_I)
    If lErro <> AD_BOOL_TRUE Then gError 116665
    
    lErro = objRelOpcoes.IncluirParametro("TADMINISTRADORAINIC", AdministradoraDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116681
    
    'Inclui parametro de AdministradoraAte
    lErro = objRelOpcoes.IncluirParametro("NADMINISTRADORAFIM", sAdmMeioPagto_F)
    If lErro <> AD_BOOL_TRUE Then gError 116666
    
    'Inclui parametro de AdministradoraAte
    lErro = objRelOpcoes.IncluirParametro("TADMINISTRADORAFIM", AdministradoraAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116682
    
    'Inclui parametro de DataDe
    lErro = objRelOpcoes.IncluirParametro("DINI", sDataInic)
    If lErro <> AD_BOOL_TRUE Then gError 116667

    'Inclui parametro de DataAte
    lErro = objRelOpcoes.IncluirParametro("DFIM", sDataFim)
    If lErro <> AD_BOOL_TRUE Then gError 116668
       
    'Inclui parametro de BorderoDe
    lErro = objRelOpcoes.IncluirParametro("NBORDEROINIC", BorderoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116669
    
    'Inclui parametro de BorderoAte
    lErro = objRelOpcoes.IncluirParametro("NBORDEROFIM", BorderoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 116670
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        'Inclui Parametro Filial Empresa
        lErro = objRelOpcoes.IncluirParametro("NFILIALEMPRESA", CStr(giFilialEmpresa))
        If lErro <> AD_BOOL_TRUE Then gError 116671
    
    End If
    
    'Aciona Rotina que monta_expressão que será usada para gerar relatório
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAdmMeioPagto_I, sAdmMeioPagto_F)
    If lErro <> SUCESSO Then gError 116672
        
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr


    Select Case gErr

        Case 116663 To 116672
        
        Case 116681, 116682
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172444)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAdmMeioPagto_I As String, sAdmMeioPagto_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    'Verifica se campo AdministradoraDe foi preenchido
    If Trim(AdministradoraDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressao o Valor de AdministradoraDe
        sExpressao = sExpressao & "Administradora >= " & Forprint_ConvLong(CLng(sAdmMeioPagto_I))
        
    End If

    'Verifica se campo AdministradoraAte foi preenchido
    If Trim(AdministradoraAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de AdministradoraAte
        sExpressao = sExpressao & "Administradora <= " & Forprint_ConvLong(CLng(sAdmMeioPagto_F))

    End If
    
    'Verifica se campo DataDe foi preenchido
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de DataDe
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If
    
    'Verifica se campo DataAte foi preenchido
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de DataAte
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172445)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim Cancel As Boolean

On Error GoTo Erro_PreencherParametrosNaTela
    'Limpa a Tela
    lErro = Limpa_Tela
    If lErro <> SUCESSO Then gError 116687
    
    'Carrega parametros do relatorio gravado
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 166673
            
    'pega parâmetro Administradora Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NADMINISTRADORAINIC", sParam)
    If lErro Then gError 166674
    
    'Preenche campo AdministradoraDe
    AdministradoraDe.Text = sParam
    
    'verifica validade de AdministradoraDe
    AdministradoraDe_Validate (Cancel)
    If Cancel = True Then gError 116685
    
    'pega parâmetro Administradora Final e exibe
    lErro = objRelOpcoes.ObterParametro("NADMINISTRADORAFIM", sParam)
    If lErro Then gError 166675
    
    'Preenche campo AdministradoraAte
    AdministradoraAte.Text = sParam
                
    'verifica validade de AdministradoraAte
    AdministradoraAte_Validate (Cancel)
    If Cancel = True Then gError 116686
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError 166676

    'Preenche campo DataDe
    Call DateParaMasked(DataDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 166677

    'Preenche campo DataAte
    Call DateParaMasked(DataAte, CDate(sParam))
        
    'Pega parametro BorderoDe e o Exibe
    lErro = objRelOpcoes.ObterParametro("NBORDEROINIC", sParam)
    If lErro <> SUCESSO Then gError 166678

    'Preenche campo BorderoDe
    BorderoDe.PromptInclude = False
    BorderoDe.Text = sParam
    BorderoDe.PromptInclude = True
    
    'Pega parametro BorderoAte e o Exibe
    lErro = objRelOpcoes.ObterParametro("NBORDEROFIM", sParam)
    If lErro <> SUCESSO Then gError 166679

    'Preenche campo BorderoAte
    BorderoAte.PromptInclude = False
    BorderoAte.Text = sParam
    BorderoAte.PromptInclude = True
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 166673 To 116679
        
        Case 116685
            AdministradoraDe.Text = ""
            
        Case 116686
            AdministradoraAte.Text = ""
            
        Case 116687
            BorderoDe.Text = ""
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172446)

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
    
   'Limpa campos de Administradora
    AdministradoraDe.Text = ""
    AdministradoraAte.Text = ""

    Exit Function

Erro_Limpa_Tela:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172447)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Bordero Ticket"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelBorderoTicket"

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

Private Sub LabelAdministradoraAte_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelAdministradoraAte, Source, X, Y)
End Sub

Private Sub LabelAdministradoraDe_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelAdministradoraDe, Source, X, Y)
End Sub

Private Sub LabelBorderoAte_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelBorderoAte, Source, X, Y)
End Sub
Private Sub LabelBorderoDe_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelBorderoDe, Source, X, Y)
End Sub
