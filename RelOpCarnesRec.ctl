VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpCarnesRec 
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   ScaleHeight     =   6390
   ScaleMode       =   0  'User
   ScaleWidth      =   6690
   Begin VB.CheckBox DetalhaCarne 
      Caption         =   "Detalhar por Carnê"
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
      Left            =   240
      TabIndex        =   11
      Top             =   6080
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   1290
      Left            =   240
      TabIndex        =   28
      Top             =   3120
      Width           =   4215
      Begin MSMask.MaskEdBox ClienteDe 
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteAte 
         Height          =   300
         Left            =   960
         TabIndex        =   6
         Top             =   825
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         Left            =   525
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   420
         Width           =   315
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
         Left            =   420
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   885
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Cliente"
      Height          =   1450
      Left            =   240
      TabIndex        =   27
      Top             =   4515
      Width           =   4215
      Begin VB.CheckBox AgrupaTipo 
         Caption         =   "Agrupar por Tipo de Cliente"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1080
         Width           =   3015
      End
      Begin VB.ComboBox ComboTipo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   623
         Width           =   2145
      End
      Begin VB.OptionButton OptionUmTipo 
         Caption         =   "Apenas do Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   630
         Width           =   1755
      End
      Begin VB.OptionButton OptionTodosTipos 
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
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   315
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpCarnesRec.ctx":0000
      Left            =   1080
      List            =   "RelOpCarnesRec.ctx":0002
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpCarnesRec.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpCarnesRec.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpCarnesRec.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpCarnesRec.ctx":081A
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
      Picture         =   "RelOpCarnesRec.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   945
      Width           =   1605
   End
   Begin VB.Frame FrameDataVencimento 
      Caption         =   "Data Vencimento"
      Height          =   735
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Width           =   4215
      Begin MSComCtl2.UpDown UpDownVencimentoDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VencimentoDe 
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
      Begin MSComCtl2.UpDown UpDownVencimentoAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VencimentoAte 
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
      Begin VB.Label LabelVencimentoDe 
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
         TabIndex        =   24
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelVencimentoAte 
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
         TabIndex        =   23
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame FrameCarne 
      Caption         =   "Carnê"
      Height          =   1290
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   4215
      Begin MSMask.MaskEdBox CarneDe 
         Height          =   300
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         Mask            =   "99999999999999999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CarneAte 
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   825
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         Mask            =   "99999999999999999999"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCarneDe 
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
         Left            =   720
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   420
         Width           =   315
      End
      Begin VB.Label LabelCarneAte 
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
         Left            =   675
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   885
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
      TabIndex        =   26
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpCarnesRec"
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

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoCarne As AdmEvento
Attribute objEventoCarne.VB_VarHelpID = -1

'Usado apenas nessa tela
Type typeCamposCarne
    sCheckTipo As String
    sCliente_I As String
    sCliente_F As String
    sCarne_I As String
    sCarne_F As String
    sClienteTipo As String
End Type

Dim giClienteDe As Integer
Dim giCarneDe As Integer

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Carnês a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpCarnesRec"
    
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

Private Sub LabelCarneDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCarneDe, Source, X, Y)
End Sub

Private Sub LabelCarneDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCarneDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCarneAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCarneDe, Source, X, Y)
End Sub

Private Sub LabelCarneAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCarneAte, Button, Shift, X, Y)
End Sub

Private Sub LabelVencimentoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVencimentoDe, Source, X, Y)
End Sub

Private Sub LabelVencimentoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVencimentoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelVencimentoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVencimentoDe, Source, X, Y)
End Sub

Private Sub LabelVencimentoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVencimentoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Private Sub VencimentoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(VencimentoDe)
End Sub

Private Sub VencimentoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(VencimentoAte)
End Sub

Private Sub VencimentoDe_Validate(Cancel As Boolean)
' valida e critica a Vencimento Inicial

Dim lgErro As Long

On Error GoTo ERRO_VencimentoDe_Validate

    If Len(VencimentoDe.ClipText) > 0 Then

        lgErro = Data_Critica(VencimentoDe.Text)
        If lgErro <> SUCESSO Then gError 117120

    End If

    Exit Sub

ERRO_VencimentoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 117120

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167505)

    End Select

    Exit Sub

End Sub

Private Sub VencimentoAte_Validate(Cancel As Boolean)
' valida e critica a Vencimento final

Dim lgErro As Long

On Error GoTo ERRO_VencimentoAte_Validate

    If Len(VencimentoAte.ClipText) > 0 Then

        lgErro = Data_Critica(VencimentoAte.Text)
        If lgErro <> SUCESSO Then gError 117121

    End If

    Exit Sub

ERRO_VencimentoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 117121

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167506)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lgErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo ERRO_BotaoExcluir_Click

    'Verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 117122

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lgErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lgErro <> SUCESSO Then gError 117123

        'Retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click
           
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

ERRO_BotaoExcluir_Click:

    Select Case gErr

        Case 117122
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 117123

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167507)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lgErro As Long

On Error GoTo ERRO_BotaoLimpar_Click

    lgErro = Limpa_Relatorio(Me)
    If lgErro <> SUCESSO Then gError 117124
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
ERRO_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 117124
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167508)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click

        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click

        ElseIf Me.ActiveControl Is CarneDe Then
            Call LabelCarneDe_Click

        ElseIf Me.ActiveControl Is CarneAte Then
            Call LabelCarneAte_Click

        End If

    End If

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lgErro As Long
Dim iResultado As Integer

On Error GoTo ERRO_BotaoGravar_Click

    'Nome da opção de Relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 117031

    lgErro = PreencherRelOp(gobjRelOpcoes)
    If lgErro <> SUCESSO Then gError 117032

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lgErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lgErro <> SUCESSO Then gError 117033

    lgErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lgErro <> SUCESSO Then gError 117034
    
    Call BotaoLimpar_Click
    
    Exit Sub

ERRO_BotaoGravar_Click:

    Select Case gErr

        Case 117031
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 117032, 117033, 117034

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167509)

    End Select

    Exit Sub

End Sub

Private Sub ClienteDe_Validate(Cancel As Boolean)

Dim lgErro As Long
Dim objCliente As New ClassCliente

On Error GoTo ERRO_ClienteDe_Validate

    If Len(Trim(ClienteDe.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lgErro = TP_Cliente_Le2(ClienteDe, objCliente, 0)
        If lgErro <> SUCESSO Then gError 117125

    End If
    
    giClienteDe = 1
    
    Exit Sub

ERRO_ClienteDe_Validate:

    Cancel = True


    Select Case gErr

        Case 117125
            'Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167510)

    End Select

End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)

Dim lgErro As Long
Dim objCliente As New ClassCliente

On Error GoTo ERRO_ClienteAte_Validate

    If Len(Trim(ClienteAte.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lgErro = TP_Cliente_Le2(ClienteAte, objCliente, 0)
        If lgErro <> SUCESSO Then gError 117126

    End If
    
    giClienteDe = 0
 
    Exit Sub

ERRO_ClienteAte_Validate:

    Cancel = True


    Select Case gErr

        Case 117126
             'Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167511)

    End Select

End Sub

Private Sub BotaoExecutar_Click()

Dim lgErro As Long

On Error GoTo ERRO_BotaoExecutar_Click

    lgErro = PreencherRelOp(gobjRelOpcoes)
    If lgErro <> SUCESSO Then gError 117127
    
     If (AgrupaTipo.Value = Unchecked) And (DetalhaCarne.Value = Unchecked) Then gobjRelatorio.sNomeTsk = "CNRC"
     If (AgrupaTipo.Value = Checked) And (DetalhaCarne.Value = Checked) Then gobjRelatorio.sNomeTsk = "CNRCTPCN"
     If (AgrupaTipo.Value = Checked) And (DetalhaCarne.Value = Unchecked) Then gobjRelatorio.sNomeTsk = "CNRCTP"
     If (AgrupaTipo.Value = Unchecked) And (DetalhaCarne.Value = Checked) Then gobjRelatorio.sNomeTsk = "CNRCCN"
        
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

ERRO_BotaoExecutar_Click:

    Select Case gErr

        Case 117127

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167512)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lgErro As Long

On Error GoTo ERRO_Form_Load

    Set objEventoCliente = New AdmEvento
        
    'Preenche com os Tipos de Clientes
    lgErro = PreencheComboTipo()
    If lgErro <> SUCESSO Then gError 117128
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

ERRO_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        Case 117128
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167513)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lgErro As Long

On Error GoTo ERRO_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 117129
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lgErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lgErro <> SUCESSO Then gError 117130
    
    Trata_Parametros = SUCESSO

    Exit Function

ERRO_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 117130
        
        Case 117129
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167514)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(tCampo As typeCamposCarne) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Tipocliente

Dim lgErro As Long

On Error GoTo ERRO_Formata_E_Critica_Parametros
       
    'critica Cliente Inicial e Final
    If ClienteDe.Text <> "" Then
        tCampo.sCliente_I = CStr(LCodigo_Extrai(ClienteDe.Text))
    Else
        tCampo.sCliente_I = ""
    End If
    
    If ClienteAte.Text <> "" Then
        tCampo.sCliente_F = CStr(LCodigo_Extrai(ClienteAte.Text))
    Else
        tCampo.sCliente_F = ""
    End If
            
    If tCampo.sCliente_I <> "" And tCampo.sCliente_F <> "" Then
        
        If CLng(tCampo.sCliente_I) > CLng(tCampo.sCliente_F) Then gError 117131
        
    End If
    
    'critica Carne Inicial e Final
    If CarneDe.Text <> "" Then
        tCampo.sCarne_I = CStr(CarneDe.ClipText)
    Else
        tCampo.sCarne_I = ""
    End If
    
    If CarneAte.Text <> "" Then
        tCampo.sCarne_F = CStr(CarneAte.ClipText)
    Else
        tCampo.sCarne_F = ""
    End If
            
    If tCampo.sCarne_I <> "" And tCampo.sCarne_F <> "" Then
        
        If CInt(tCampo.sCarne_I) > CInt(tCampo.sCarne_F) Then gError 117132
        
    End If
            
    'Se a opção para todos os Clientes estiver selecionada
    If OptionTodosTipos.Value = True Then
        tCampo.sCheckTipo = "Todos"
        tCampo.sClienteTipo = ""
    
    'Se a opção para apenas um Cliente estiver selecionada
    Else
        'TEm que indicar o tipo do Cliente
        If ComboTipo.Text = "" Then gError 117133
        tCampo.sCheckTipo = "Um"
        tCampo.sClienteTipo = ComboTipo.Text
        
    End If
        
    'data vencimento inicial nao pode ser maior que a final
    If Trim(VencimentoDe.ClipText) <> "" And Trim(VencimentoAte.ClipText) <> "" Then
    
        If CDate(VencimentoDe.Text) > CDate(VencimentoAte.Text) Then gError 117134
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

ERRO_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 117131
            lgErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteDe.SetFocus
            
        Case 117132
            lgErro = Rotina_Erro(vbOKOnly, "ERRO_CARNE_INICIAL_MAIOR", gErr)
            CarneDe.SetFocus
                
        Case 117133
            lgErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_PREENCHIDO", gErr)
            ComboTipo.SetFocus
                                
        Case 117134
            lgErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCTO_INICIAL_MAIOR", gErr)
            VencimentoDe.SetFocus
            
        Case Else
            lgErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167515)

    End Select

    Exit Function

End Function

Private Sub UpDownVencimentoDe_DownClick()

Dim lgErro As Long

On Error GoTo ERRO_UpDownVencimentoDe_DownClick

    lgErro = Data_Up_Down_Click(VencimentoDe, DIMINUI_DATA)
    If lgErro <> SUCESSO Then gError 117135

    Exit Sub

ERRO_UpDownVencimentoDe_DownClick:

    Select Case gErr

        Case 117135
            VencimentoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167516)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimentoDe_UpClick()

Dim lgErro As Long

On Error GoTo ERRO_UpDownVencimentoDe_UpClick

    lgErro = Data_Up_Down_Click(VencimentoDe, AUMENTA_DATA)
    If lgErro <> SUCESSO Then gError 117136

    Exit Sub

ERRO_UpDownVencimentoDe_UpClick:

    Select Case gErr

        Case 117136
            VencimentoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167517)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimentoAte_DownClick()

Dim lgErro As Long

On Error GoTo ERRO_UpDownVencimentoAte_DownClick

    lgErro = Data_Up_Down_Click(VencimentoAte, DIMINUI_DATA)
    If lgErro <> SUCESSO Then gError 117137

    Exit Sub

ERRO_UpDownVencimentoAte_DownClick:

    Select Case gErr

        Case 117137
            VencimentoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167518)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimentoAte_UpClick()

Dim lgErro As Long

On Error GoTo ERRO_UpDownVencimentoAte_UpClick

    lgErro = Data_Up_Down_Click(VencimentoAte, AUMENTA_DATA)
    If lgErro <> SUCESSO Then gError 117138

    Exit Sub

ERRO_UpDownVencimentoAte_UpClick:

    Select Case gErr

        Case 117138
            VencimentoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167519)

    End Select

    Exit Sub

End Sub

Private Sub OptionTodosTipos_Click()

On Error GoTo ERRO_OptionTodosTipos_Click
    
    'Limpa e desabilita a ComboTipo
    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = False
    'OptionTodosTipos.Value = True
    
    Exit Sub

ERRO_OptionTodosTipos_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167520)

    End Select

    Exit Sub
    
End Sub

Function PreencheComboTipo() As Long

Dim lgErro As Long
Dim colCodNomeTPCliente As New AdmColCodigoNome
Dim objCodNomeTPCliente As New AdmCodigoNome

On Error GoTo ERRO_PreencheComboTipo
    
    'Preenche a Colecao com os Tipos de clientes
    lgErro = CF("Cod_Nomes_Le", "TiposdeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodNomeTPCliente)
    If lgErro <> SUCESSO Then gError 117139
    
   'preenche a ListBox ComboTipo com os objetos da colecao
    For Each objCodNomeTPCliente In colCodNomeTPCliente
        ComboTipo.AddItem objCodNomeTPCliente.iCodigo & SEPARADOR & objCodNomeTPCliente.sNome
        ComboTipo.ItemData(ComboTipo.NewIndex) = objCodNomeTPCliente.iCodigo
    Next
        
    PreencheComboTipo = SUCESSO

    Exit Function
    
ERRO_PreencheComboTipo:

    PreencheComboTipo = gErr

    Select Case gErr

    Case 117139
    
    Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167521)

    End Select

    Exit Function

End Function

Private Sub OptionUmTipo_Click()

On Error GoTo ERRO_OptionUmTipo_Click
    
    'Limpa Combo Tipo e a Habilita
    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = True
    
    Exit Sub

ERRO_OptionUmTipo_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167522)

    End Select

    Exit Sub
    
End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lgErro As Long
Dim tCampo As typeCamposCarne

On Error GoTo ERRO_PreencherRelOp
            
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lgErro = Formata_E_Critica_Parametros(tCampo)
    If lgErro <> SUCESSO Then gError 117140

    lgErro = objRelOpcoes.Limpar
    If lgErro <> AD_BOOL_TRUE Then gError 117141
         
    'Preenche o Cliente Inicial
    lgErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", tCampo.sCliente_I)
    If lgErro <> AD_BOOL_TRUE Then gError 117142
    
    lgErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", Trim(ClienteDe.Text))
    If lgErro <> AD_BOOL_TRUE Then gError 117143
    
    'Preenche o Cliente Final
    lgErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", tCampo.sCliente_F)
    If lgErro <> AD_BOOL_TRUE Then gError 117144
     
    lgErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", Trim(ClienteAte.Text))
    If lgErro <> AD_BOOL_TRUE Then gError 117145
                   
    'Preenche o Código do Carnê Inicial
    lgErro = objRelOpcoes.IncluirParametro("NCARNEINIC", tCampo.sCarne_I)
    If lgErro <> AD_BOOL_TRUE Then gError 117146
    
    'Preenche o Código do Carnê Final
    lgErro = objRelOpcoes.IncluirParametro("NCARNEFIM", tCampo.sCarne_F)
    If lgErro <> AD_BOOL_TRUE Then gError 117147
    
    'Preenche o tipo do Cliente
    lgErro = objRelOpcoes.IncluirParametro("TTIPOCLIENTE", tCampo.sClienteTipo)
    If lgErro <> AD_BOOL_TRUE Then gError 117148
    
    'Preenche com a Opcao Tipocliente(TodosClientes ou um Cliente)
    lgErro = objRelOpcoes.IncluirParametro("TOPTIPO", tCampo.sCheckTipo)
    If lgErro <> AD_BOOL_TRUE Then gError 117149
                  
    'Preenche com o Detalhar por carnê
    lgErro = objRelOpcoes.IncluirParametro("NDETCARN", CStr(DetalhaCarne.Value))
    If lgErro <> AD_BOOL_TRUE Then gError 117150
    
    'Preenche com o Agrupa por tipo
    lgErro = objRelOpcoes.IncluirParametro("NAGRUPTIPO", CStr(AgrupaTipo.Value))
    If lgErro <> AD_BOOL_TRUE Then gError 117151
    
    If VencimentoDe.ClipText <> "" Then
        lgErro = objRelOpcoes.IncluirParametro("DVENINIC", VencimentoDe.Text)
    Else
        lgErro = objRelOpcoes.IncluirParametro("DVENINIC", CStr(DATA_NULA))
    End If
    If lgErro <> AD_BOOL_TRUE Then gError 117152

    If VencimentoAte.ClipText <> "" Then
        lgErro = objRelOpcoes.IncluirParametro("DVENFIM", VencimentoAte.Text)
    Else
        lgErro = objRelOpcoes.IncluirParametro("DVENFIM", CStr(DATA_NULA))
    End If
    If lgErro <> AD_BOOL_TRUE Then gError 117153
    
    'Faz a selecao
    lgErro = Monta_Expressao_Selecao(objRelOpcoes, tCampo)
    If lgErro <> SUCESSO Then gError 117154

    PreencherRelOp = SUCESSO

    Exit Function

ERRO_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 117140 To 117154
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167523)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lgErro As Long
Dim sParam As String
Dim sTipoCliente As String
Dim iIndice As Integer

On Error GoTo ERRO_PreencherParametrosNaTela

    lgErro = objRelOpcoes.Carregar
    If lgErro <> SUCESSO Then gError 117155
   
    'pega Cliente inicial e exibe
    lgErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lgErro <> SUCESSO Then gError 117156
    
    ClienteDe.Text = sParam
    Call ClienteDe_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lgErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lgErro <> SUCESSO Then gError 117158
    
    ClienteAte.Text = sParam
    Call ClienteAte_Validate(bSGECancelDummy)
    
    'pega Carnê inicial e exibe
    lgErro = objRelOpcoes.ObterParametro("NCARNEINIC", sParam)
    If lgErro <> SUCESSO Then gError 117157
    
    CarneDe.PromptInclude = False
    CarneDe.Text = sParam
    CarneDe.PromptInclude = True
    
    'pega  Carnê final e exibe
    lgErro = objRelOpcoes.ObterParametro("NCARNEFIM", sParam)
    If lgErro <> SUCESSO Then gError 117159
    
    CarneAte.PromptInclude = False
    CarneAte.Text = sParam
    CarneAte.PromptInclude = True
                
    'pega  Tipo cliente e Exibe
    lgErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lgErro <> SUCESSO Then gError 117160
                   
    If sParam = "Todos" Then
    
        Call OptionTodosTipos_Click
    
    Else
        'se é "um tipo só" então exibe o tipo
        lgErro = objRelOpcoes.ObterParametro("TTIPOCLIENTE", sTipoCliente)
        If lgErro <> SUCESSO Then gError 117161
                            
        OptionUmTipo.Value = True
        ComboTipo.Enabled = True
        
        If sTipoCliente = "" Then
            ComboTipo.ListIndex = -1
        Else
            ComboTipo.Text = sTipoCliente
        End If
    End If
    
    lgErro = objRelOpcoes.ObterParametro("NDETCARN", sParam)
    If lgErro <> SUCESSO Then gError 117162
    
    DetalhaCarne.Value = CInt(sParam)
        
    lgErro = objRelOpcoes.ObterParametro("NAGRUPTIPO", sParam)
    If lgErro <> SUCESSO Then gError 117163
    
    AgrupaTipo.Value = CInt(sParam)
          
    'pega data vencimento inicial e exibe
    lgErro = objRelOpcoes.ObterParametro("DVENINIC", sParam)
    If lgErro <> SUCESSO Then gError 117164

    Call DateParaMasked(VencimentoDe, CDate(sParam))
    
    'pega data vencimento final e exibe
    lgErro = objRelOpcoes.ObterParametro("DVENFIM", sParam)
    If lgErro <> SUCESSO Then gError 117165

    Call DateParaMasked(VencimentoAte, CDate(sParam))
               
    PreencherParametrosNaTela = SUCESSO

    Exit Function

ERRO_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 117155 To 11765
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167524)

    End Select

    Exit Function

End Function

Private Sub ClienteAte_LostFocus()
    Call ClienteAte_Validate(bSGECancelDummy)
End Sub

Private Sub ClienteDe_LostFocus()
    Call ClienteDe_Validate(bSGECancelDummy)
End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, tCampo As typeCamposCarne) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lgErro As Long

On Error GoTo ERRO_Monta_Expressao_Selecao

   If tCampo.sCliente_I <> "" Then sExpressao = "Cliente >= " & Forprint_ConvInt(CInt(tCampo.sCliente_I))

   If tCampo.sCliente_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvInt(CInt(tCampo.sCliente_F))

    End If
    
   If tCampo.sCarne_I <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Carne >= " & Forprint_ConvInt(CInt(tCampo.sCarne_I))

   End If
    
   If tCampo.sCarne_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Carne <= " & Forprint_ConvInt(CInt(tCampo.sCarne_F))

    End If
                 
    'Se a opção para apenas um cliente estiver selecionada
    If tCampo.sCheckTipo = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoCliente = " & Forprint_ConvInt(CInt(Codigo_Extrai(tCampo.sClienteTipo)))

    End If
    
    If Trim(VencimentoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vencimento >= " & Forprint_ConvData(CDate(VencimentoDe.Text))

    End If
    
    If Trim(VencimentoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vencimento <= " & Forprint_ConvData(CDate(VencimentoAte.Text))

    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(CInt(giFilialEmpresa))
    End If
    
        'passa a expressão completa para o obj
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

ERRO_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167525)

    End Select

    Exit Function

End Function

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    If giClienteDe = 1 Then
        ClienteDe.PromptInclude = False
        ClienteDe.Text = CStr(objCliente.lCodigo)
        ClienteDe.PromptInclude = True
        ClienteDe_Validate (bSGECancelDummy)
    Else
        ClienteAte.PromptInclude = False
        ClienteAte.Text = CStr(objCliente.lCodigo)
        ClienteAte.PromptInclude = True
        ClienteAte_Validate (bSGECancelDummy)
    End If

    Me.Show

     Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteDe = 0
    
    If Len(Trim(ClienteAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteAte.Text)
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteDe = 1

    If Len(Trim(ClienteDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteDe.Text)
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub


Private Sub LabelCarneDe_Click()

Dim objCarne As New ClassCarne
Dim colSelecao As Collection

    giCarneDe = 1

    If Len(Trim(CarneDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCarne.lNumIntDoc = CarneDe.ClipText
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("CarneLista", colSelecao, objCarne, objEventoCarne)

End Sub

Private Sub LabelCarneAte_Click()

Dim objCarne As New ClassCarne
Dim colSelecao As Collection

    giCarneDe = 0

    If Len(Trim(CarneAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCarne.lNumIntDoc = CarneAte.ClipText
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("CarneLista", colSelecao, objCarne, objEventoCarne)

End Sub

Private Sub objEventoCarne_evSelecao(obj1 As Object)

Dim objCarne As ClassCarne
    
    Set objCarne = obj1
    
    'Preenche campo Cliente
    If giCarneDe = 1 Then
        CarneDe.PromptInclude = False
        CarneDe.Text = CStr(objCarne.lCupomFiscal)
        CarneDe.PromptInclude = True
        
    Else
        CarneAte.PromptInclude = False
        CarneAte.Text = CStr(objCarne.lCupomFiscal)
        CarneAte.PromptInclude = True
    End If

    Me.Show

     Exit Sub

End Sub
