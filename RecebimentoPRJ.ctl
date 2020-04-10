VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RecebimentoPRJ 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.CheckBox Checkbox_Verifica_Sintaxe 
      Caption         =   "Verifica Sintaxe ao Sair do Campo"
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
      Left            =   135
      TabIndex        =   8
      Top             =   1395
      Value           =   1  'Checked
      Width           =   3285
   End
   Begin VB.CheckBox CronFisFin 
      Caption         =   "Inclui no Cronograma Fís. Fin."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3570
      TabIndex        =   9
      Top             =   1365
      Width           =   2910
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6795
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   29
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RecebimentoPRJ.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RecebimentoPRJ.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RecebimentoPRJ.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RecebimentoPRJ.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condições"
      Height          =   4260
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   8850
      Begin VB.TextBox Regra 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3960
         TabIndex        =   37
         Top             =   615
         Width           =   3240
      End
      Begin VB.TextBox Observacao 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1110
         TabIndex        =   36
         Top             =   2085
         Width           =   2505
      End
      Begin VB.ComboBox CondPagto 
         Height          =   315
         Left            =   4860
         TabIndex        =   34
         Top             =   1320
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Percentual 
         Height          =   270
         Left            =   4845
         TabIndex        =   35
         Top             =   855
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin VB.TextBox Descricao 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   540
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   3585
         Width           =   8565
      End
      Begin VB.ComboBox Mnemonicos 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "RecebimentoPRJ.ctx":0994
         Left            =   135
         List            =   "RecebimentoPRJ.ctx":09A1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3195
         Width           =   2520
      End
      Begin VB.ComboBox Funcoes 
         Height          =   315
         Left            =   5055
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3195
         Width           =   2415
      End
      Begin VB.ComboBox Operadores 
         Height          =   315
         Left            =   7635
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3195
         Width           =   1050
      End
      Begin VB.ComboBox Etapa 
         Height          =   315
         ItemData        =   "RecebimentoPRJ.ctx":09C3
         Left            =   2760
         List            =   "RecebimentoPRJ.ctx":09C5
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3195
         Width           =   2220
      End
      Begin MSFlexGridLib.MSFlexGrid GridRegras 
         Height          =   2550
         Left            =   105
         TabIndex        =   11
         Top             =   195
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   4498
         _Version        =   393216
         Rows            =   50
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Etapas:"
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
         Left            =   2760
         TabIndex        =   28
         Top             =   2955
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Funções:"
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
         Left            =   5055
         TabIndex        =   27
         Top             =   2970
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mnemônicos:"
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
         Left            =   135
         TabIndex        =   26
         Top             =   2940
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Operadores:"
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
         Left            =   7665
         TabIndex        =   25
         Top             =   2955
         Width           =   1050
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1860
      Picture         =   "RecebimentoPRJ.ctx":09C7
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Numeração Automática"
      Top             =   555
      Width           =   300
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4515
      TabIndex        =   6
      Top             =   915
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   900
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   300
      Left            =   1095
      TabIndex        =   2
      Top             =   540
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   4500
      TabIndex        =   4
      Top             =   525
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Projeto 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzidoPRJ 
      Height          =   315
      Left            =   4500
      TabIndex        =   1
      Top             =   120
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Contrato 
      Height          =   300
      Left            =   7590
      TabIndex        =   10
      Top             =   1350
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Proposta 
      Height          =   300
      Left            =   7590
      TabIndex        =   7
      Top             =   930
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label LabelContrato 
      AutoSize        =   -1  'True
      Caption         =   "Contrato:"
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
      Height          =   195
      Left            =   6720
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   38
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label LabelProposta 
      AutoSize        =   -1  'True
      Caption         =   "Proposta:"
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
      Height          =   195
      Left            =   6720
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   33
      Top             =   975
      Width           =   825
   End
   Begin VB.Label LabelNomeRedPRJ 
      Caption         =   "Nome Projeto:"
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
      Height          =   315
      Left            =   3240
      TabIndex        =   32
      Top             =   165
      Width           =   1275
   End
   Begin VB.Label LabelProjeto 
      AutoSize        =   -1  'True
      Caption         =   "Projeto:"
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
      Left            =   360
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   31
      Top             =   165
      Width           =   675
   End
   Begin VB.Label Label17 
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
      Left            =   3930
      TabIndex        =   24
      Top             =   555
      Width           =   510
   End
   Begin VB.Label LabelNumero 
      AutoSize        =   -1  'True
      Caption         =   "Número:"
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
      Left            =   315
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   22
      Top             =   570
      Width           =   720
   End
   Begin VB.Label ClienteLabel 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   21
      Top             =   960
      Width           =   660
   End
   Begin VB.Label LabelFilial 
      AutoSize        =   -1  'True
      Caption         =   " Filial:"
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
      Left            =   3930
      TabIndex        =   20
      Top             =   960
      Width           =   525
   End
End
Attribute VB_Name = "RecebimentoPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Const KEYCODE_VERIFICAR_SINTAXE = vbKeyF5

Dim objGridRegra As AdmGrid
Dim iGrid_Regra_Col As Integer
Dim iGrid_Percentual_Col As Integer
Dim iGrid_CondPagto_Col As Integer
Dim iGrid_Observacao_Col As Integer

Dim iAlterado As Integer

Dim sProjetoAnt As String
Dim sNomeProjetoAnt As String

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoPRJ As AdmEvento
Attribute objEventoPRJ.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProposta As AdmEvento
Attribute objEventoProposta.VB_VarHelpID = -1
Private WithEvents objEventoContrato As AdmEvento
Attribute objEventoContrato.VB_VarHelpID = -1

Const CONTABILIZACAO_OBRIGATORIA = 1
Const CONTABILIZACAO_NAO_OBRIGATORIA = 0

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Recebimentos relacionados a Projeto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RecebimentoPRJ"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        ElseIf Me.ActiveControl Is NomeReduzidoPRJ Then
            Call LabelNomeRedPRJ_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call LabelNumero_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        End If
    
    ElseIf KeyCode = KEYCODE_VERIFICAR_SINTAXE Then
        If Checkbox_Verifica_Sintaxe.Value = MARCADO Then
            Checkbox_Verifica_Sintaxe.Value = DESMARCADO
        Else
            Checkbox_Verifica_Sintaxe.Value = MARCADO
        End If
    End If
End Sub

Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

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
'********************************************************

Public Sub Regra_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Regra_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegra)
End Sub

Public Sub Regra_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegra)
End Sub

Public Sub Regra_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridRegra.objControle = Regra
    lErro = Grid_Campo_Libera_Foco(objGridRegra)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Percentual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Percentual_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegra)
End Sub

Public Sub Percentual_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegra)
End Sub

Public Sub Percentual_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridRegra.objControle = Percentual
    lErro = Grid_Campo_Libera_Foco(objGridRegra)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub CondPagto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CondPagto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegra)
End Sub

Public Sub CondPagto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegra)
End Sub

Public Sub CondPagto_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridRegra.objControle = CondPagto
    lErro = Grid_Campo_Libera_Foco(objGridRegra)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Observacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegra)
End Sub

Public Sub Observacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegra)
End Sub

Public Sub Observacao_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridRegra.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridRegra)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoNumero = Nothing
    Set objEventoPRJ = Nothing
    Set objEventoCliente = Nothing
    Set objEventoProposta = Nothing
    Set objEventoContrato = Nothing
    
    Set objGridRegra = Nothing
    
End Sub

Public Sub Funcoes_Click()

Dim iPos As Integer
Dim lErro As Long
Dim objFormulaFuncao As New ClassFormulaFuncao
Dim lPos As Long
Dim sFuncao As String
    
On Error GoTo Erro_Funcoes_Click
    
    objFormulaFuncao.sFuncaoCombo = Funcoes.Text
    
    'retorna os dados da funcao passada como parametro
    lErro = CF("FormulaFuncao_Le", objFormulaFuncao)
    If lErro <> SUCESSO And lErro <> 36088 Then gError 185608
    
    Descricao.Text = objFormulaFuncao.sFuncaoDesc
    
    lPos = InStr(1, Funcoes.Text, "(")
    If lPos = 0 Then
        sFuncao = Funcoes.Text
    Else
        sFuncao = Mid(Funcoes.Text, 1, lPos)
    End If
    
    lErro = Funcoes1(sFuncao)
    If lErro <> SUCESSO Then gError 185609
    
    Exit Sub
    
Erro_Funcoes_Click:

    Select Case gErr
    
        Case 185608, 185609
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185610)
            
    End Select
        
    Exit Sub

End Sub

Public Sub GridRegras_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridRegra, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegra, iAlterado)
    End If
    
End Sub

Public Sub GridRegras_GotFocus()
    
    Call Grid_Recebe_Foco(objGridRegra)

End Sub

Public Sub GridRegras_EnterCell()
    
    Call Grid_Entrada_Celula(objGridRegra, iAlterado)
    
End Sub

Public Sub GridRegras_LeaveCell()
    
    Call Saida_Celula(objGridRegra)
    
End Sub

Public Sub GridRegras_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridRegra)
    
End Sub

Public Sub GridRegras_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRegra, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegra, iAlterado)
    End If

End Sub

Public Sub GridRegras_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridRegra)

End Sub

Public Sub GridRegras_RowColChange()
    Call Grid_RowColChange(objGridRegra)
      
End Sub

Public Sub GridRegras_Scroll()

    Call Grid_Scroll(objGridRegra)
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objGridRegra = New AdmGrid
    
    Set objEventoNumero = New AdmEvento
    Set objEventoPRJ = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoProposta = New AdmEvento
    Set objEventoContrato = New AdmEvento
    
    'inicializa o grid de lancamentos padrão
    lErro = Inicializa_Grid_Regras(objGridRegra)
    If lErro <> SUCESSO Then gError 185611

    'carrega a combobox de funcoes
    lErro = Carga_Combobox_Funcoes()
    If lErro <> SUCESSO Then gError 185612
    
    'carrega a combobox de operadores
    lErro = Carga_Combobox_Operadores()
    If lErro <> SUCESSO Then gError 185613
    
    'Carrega os mnemônicos de projetos
    lErro = Carga_Combobox_Mnemonicos
    If lErro <> SUCESSO Then gError 185614
    
    'Carrega as condições de pagamento
    lErro = Carrega_CondicaoPagamento(CondPagto)
    If lErro <> SUCESSO Then gError 185691
    
    lErro = Inicializa_Mascara_Projeto(Projeto)
    If lErro <> SUCESSO Then gError 189067
       
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 185612 To 185614, 185691, 189067
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185615)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional ByVal objPRJReceb As ClassPRJRecebPagto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objPRJReceb Is Nothing) Then
    
        objPRJReceb.iFilialEmpresa = giFilialEmpresa
        objPRJReceb.iTipo = PRJ_TIPO_RECEB
    
        lErro = Traz_Recebimento_Tela(objPRJReceb)
        If lErro <> SUCESSO Then gError 185731
    
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 185731
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185616)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Function Traz_Recebimento_Tela(ByVal objPRJReceb As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim objPRJPagtoRegras As ClassPRJRecebPagtoRegras
Dim objProposta As New ClassPRJPropostas
Dim objContrato As New ClassPRJContratos
Dim iLinha As Integer
Dim lErroAux As Long

On Error GoTo Erro_Traz_Recebimento_Tela

    Call Limpa_Tela_RecebPRJ

    'Lê a Etapa que está sendo Passada
    lErroAux = CF("PRJRecebPagto_Le", objPRJReceb)
    If lErroAux <> SUCESSO And lErroAux <> ERRO_LEITURA_SEM_DADOS Then gError 185687

    If objPRJReceb.lNumero <> 0 Then
        Numero.PromptInclude = False
        Numero.Text = objPRJReceb.lNumero
        Numero.PromptInclude = True
    End If
    If objPRJReceb.dValor <> 0 Then Valor.Text = Format(objPRJReceb.dValor, "STANDARD")
    
    If objPRJReceb.lCliForn <> 0 Then
        Cliente.Text = objPRJReceb.lCliForn
        Call Cliente_Validate(bSGECancelDummy)
    End If
    
    Call Combo_Seleciona(Filial, objPRJReceb.iFilial)
    
    If objPRJReceb.lNumIntDocPRJ <> 0 Then

        objProjeto.lNumIntDoc = objPRJReceb.lNumIntDocPRJ
        
        'Lê o Projetos que está sendo Passado
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185688
        
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189116
        
        Call Projeto_Validate(bSGECancelDummy)
        
    End If
        
    If objPRJReceb.lNumIntDocProposta <> 0 Then
    
        objProposta.lNumIntDoc = objPRJReceb.lNumIntDocProposta
    
        lErro = CF("PRJPropostas_Le_NumIntDoc", objProposta)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189004
        
        Proposta.Text = objProposta.sCodigo
        Call Proposta_Validate(bSGECancelDummy)
    
    End If
    
    If objPRJReceb.lNumIntDocContrato <> 0 Then
    
        objContrato.lNumIntDoc = objPRJReceb.lNumIntDocContrato
    
        lErro = CF("PRJContratos_Le_NumIntDoc", objContrato)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189004
        
        Contrato.Text = objContrato.sCodigo
        Call Contrato_Validate(bSGECancelDummy)
    
    End If

    If lErroAux = SUCESSO Then
        
        If objPRJReceb.iIncluiCFF = MARCADO Then
            CronFisFin.Value = vbChecked
        Else
            CronFisFin.Value = vbUnchecked
        End If
        
        iLinha = 0
        For Each objPRJPagtoRegras In objPRJReceb.colRegras
        
            iLinha = iLinha + 1
            GridRegras.TextMatrix(iLinha, iGrid_Regra_Col) = objPRJPagtoRegras.sRegra
            GridRegras.TextMatrix(iLinha, iGrid_Percentual_Col) = Format(objPRJPagtoRegras.dPercentual, "PERCENT")
            GridRegras.TextMatrix(iLinha, iGrid_Observacao_Col) = objPRJPagtoRegras.sObservacao
    
            If objPRJPagtoRegras.iCondPagto <> 0 Then
                CondPagto.Text = objPRJPagtoRegras.iCondPagto
                lErro = Combo_Seleciona_Grid(CondPagto, objPRJPagtoRegras.iCondPagto)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 185692
                GridRegras.TextMatrix(iLinha, iGrid_CondPagto_Col) = CondPagto.Text
            End If
            
        Next
    
        objGridRegra.iLinhasExistentes = objPRJReceb.colRegras.Count
        
    End If
    
    Traz_Recebimento_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Recebimento_Tela:

    Traz_Recebimento_Tela = gErr

    Select Case gErr
    
        Case 185687 To 185688, 185692, 189004, 189116
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185616)
    
    End Select
    
    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objPRJReceb As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim objPRJPagtoRegra As ClassPRJRecebPagtoRegras
Dim objCliente As New ClassCliente
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim objProposta As New ClassPRJPropostas
Dim objContrato As New ClassPRJContratos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189079

    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa
    
    'Le
    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185728
    
    'Se não encontrou => Erro
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185729
    
    objPRJReceb.lNumIntDocPRJ = objProjeto.lNumIntDoc

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 129491

        If lErro = SUCESSO Then objPRJReceb.lCliForn = objCliente.lCodigo
                            
    End If
    
    objPRJReceb.iTipo = PRJ_TIPO_RECEB
    objPRJReceb.iFilialEmpresa = giFilialEmpresa
    objPRJReceb.iFilial = Codigo_Extrai(Filial.Text)
    objPRJReceb.dValor = StrParaDbl(Valor.Text)
    objPRJReceb.lNumero = StrParaLong(Numero.Text)
    
    If CronFisFin.Value = vbChecked Then
        objPRJReceb.iIncluiCFF = MARCADO
    Else
        objPRJReceb.iIncluiCFF = DESMARCADO
    End If
    
    If Len(Trim(Proposta.Text)) > 0 Then
    
        objProposta.lNumIntDocPRJ = objProjeto.lNumIntDoc
        objProposta.sCodigo = Proposta.Text
        
        lErro = CF("PRJPropostas_Le", objProposta, False, False)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187996
        
        objPRJReceb.lNumIntDocProposta = objProposta.lNumIntDoc
        
        If objPRJReceb.lCliForn <> objProposta.lCliente Then gError 189010
        If objPRJReceb.iFilial <> objProposta.iFilialCliente Then gError 189011
        If Abs(objPRJReceb.dValor - objProposta.dValorTotal) > DELTA_VALORMONETARIO Then gError 189012
    
    End If
    
    If Len(Trim(Contrato.Text)) > 0 Then
    
        objContrato.lNumIntDocPRJ = objProjeto.lNumIntDoc
        objContrato.sCodigo = Contrato.Text
        
        lErro = CF("PRJContratos_Le", objContrato, False, False)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189437
        
        objPRJReceb.lNumIntDocContrato = objContrato.lNumIntDoc
        
        If objPRJReceb.lCliForn <> objContrato.lCliente Then gError 189438
        If objPRJReceb.iFilial <> objContrato.iFilialCliente Then gError 189439
        If Abs(objPRJReceb.dValor - objContrato.dValorTotal) > DELTA_VALORMONETARIO Then gError 189440
        If objPRJReceb.lNumIntDocProposta <> objContrato.lNumIntDocProposta Then gError 189441
    
    End If

    For iIndice = 1 To objGridRegra.iLinhasExistentes
    
        Set objPRJPagtoRegra = New ClassPRJRecebPagtoRegras
        
        objPRJPagtoRegra.sObservacao = GridRegras.TextMatrix(iIndice, iGrid_Observacao_Col)
        objPRJPagtoRegra.sRegra = GridRegras.TextMatrix(iIndice, iGrid_Regra_Col)
        objPRJPagtoRegra.iCondPagto = Codigo_Extrai(GridRegras.TextMatrix(iIndice, iGrid_CondPagto_Col))
        objPRJPagtoRegra.dPercentual = StrParaDbl(Val(GridRegras.TextMatrix(iIndice, iGrid_Percentual_Col)) / 100)
        
        objPRJReceb.colRegras.Add objPRJPagtoRegra
    
    Next
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 185728, 187996, 189079, 189437
    
        Case 185729
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
    
        Case 189010
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEB_CLIENTE_DIF_PROPOSTA", gErr)
        
        Case 189011
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEB_FILIALCLIENTE_DIF_PROPOSTA", gErr)
        
        Case 189012
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEB_VALOR_DIF_PROPOSTA", gErr)
    
        Case 189438
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEB_CLIENTE_DIF_CONTRATO", gErr)
        
        Case 189439
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEB_FILIALCLIENTE_DIF_CONTRATO", gErr)
        
        Case 189440
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEB_VALOR_DIF_CONTRATO", gErr)
        
        Case 189441
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEB_PROPOSTA_DIF_CONTRATO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185684)
    
    End Select
    
    Exit Function

End Function

Function Critica_Dados(ByVal objPRJReceb As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim objPRJPagtoRegra As ClassPRJRecebPagtoRegras
Dim iLinha As Integer
Dim dPercent As Double
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Critica_Dados

    lErro = CF("MnemonicoPRJ_Le", colMnemonico)
    If lErro <> SUCESSO Then gError 189370

    iLinha = 0
    For Each objPRJPagtoRegra In objPRJReceb.colRegras
    
        iLinha = iLinha + 1
        If Len(Trim(objPRJPagtoRegra.sRegra)) = 0 Then gError 185710
        If objPRJPagtoRegra.iCondPagto = 0 Then gError 185711
        If objPRJPagtoRegra.dPercentual = 0 Then gError 185712
        
        dPercent = dPercent + objPRJPagtoRegra.dPercentual
        
        lErro = CF("Valida_Formula_WFW", objPRJPagtoRegra.sRegra, TIPO_DATA, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 189371
    
    Next
    
    If Abs(dPercent - 1) > QTDE_ESTOQUE_DELTA2 Then gError 189015
    
    Critica_Dados = SUCESSO
    
    Exit Function

Erro_Critica_Dados:

    Critica_Dados = gErr

    Select Case gErr
    
        Case 185710
            Call Rotina_Erro(vbOKOnly, "ERRO_REGRAWFW_NAO_PREENCHIDA", gErr, iLinha)

        Case 185711
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDPAGTO_GRID_NAO_PREENCHIDA", gErr, iLinha)
        
        Case 185712
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENT_NAO_INFORMADO", gErr, iLinha)
    
        Case 189015
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENT_TOTAL_NAO_100PERC", gErr)
            
        Case 189370, 189371
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185713)
    
    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Funcoes() As Long
'carrega a combobox que contem as funcoes disponiveis

Dim lErro As Long
Dim colFormulaFuncao As New Collection
Dim objFormulaFuncao As ClassFormulaFuncao
    
On Error GoTo Erro_Carga_Combobox_Funcoes
        
    'leitura das funcoes no BD
    lErro = CF("FormulaFuncao_Le_Todos", colFormulaFuncao)
    If lErro <> SUCESSO Then gError 185617
    
    For Each objFormulaFuncao In colFormulaFuncao
        Funcoes.AddItem objFormulaFuncao.sFuncaoCombo
    Next
    
    Carga_Combobox_Funcoes = SUCESSO

    Exit Function

Erro_Carga_Combobox_Funcoes:

    Carga_Combobox_Funcoes = gErr

    Select Case gErr

        Case 185617
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185618)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Operadores() As Long
'carrega a combobox que contem os operadores disponiveis

Dim lErro As Long
Dim colFormulaOperador As New Collection
Dim objFormulaOperador As ClassFormulaOperador
    
On Error GoTo Erro_Carga_Combobox_Operadores
        
    'leitura dos operadores no BD
    lErro = CF("FormulaOperador_Le_Todos", colFormulaOperador)
    If lErro <> SUCESSO Then gError 185619
    
    For Each objFormulaOperador In colFormulaOperador
        Operadores.AddItem objFormulaOperador.sOperadorCombo
    Next
    
    Carga_Combobox_Operadores = SUCESSO

    Exit Function

Erro_Carga_Combobox_Operadores:

    Carga_Combobox_Operadores = gErr

    Select Case gErr

        Case 185619
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185620)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Mnemonicos() As Long
'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.

Dim colMnemonico As New Collection
Dim objMnemonico As ClassMnemonicoPRJ
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Mnemonicos
        
    Mnemonicos.Enabled = True
    Mnemonicos.Clear
    
    'leitura dos mnemonicos no BD
    lErro = CF("MnemonicoPRJ_Le", colMnemonico)
    If lErro <> SUCESSO Then gError 185621

    For Each objMnemonico In colMnemonico
        Mnemonicos.AddItem objMnemonico.sMnemonicoCombo
    Next
    
    Carga_Combobox_Mnemonicos = SUCESSO

    Exit Function

Erro_Carga_Combobox_Mnemonicos:

    Carga_Combobox_Mnemonicos = gErr

    Select Case gErr

        Case 185621
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185622)

    End Select
    
    Exit Function

End Function

Private Function Inicializa_Grid_Regras(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Regras
    
    'tela em questão
    Set objGridRegra.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Regra")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Cond.Pagto")
    objGridInt.colColuna.Add ("Observação")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Regra.Name)
    objGridInt.colCampo.Add (Percentual.Name)
    objGridInt.colCampo.Add (CondPagto.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    
    iGrid_Regra_Col = 1
    iGrid_Percentual_Col = 2
    iGrid_CondPagto_Col = 3
    iGrid_Observacao_Col = 4
        
    objGridInt.objGrid = GridRegras
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 10
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    GridRegras.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Regras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Regras:

    Inicializa_Grid_Regras = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185623)
        
    End Select

    Exit Function
        
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        Select Case GridRegras.Col
    
            Case iGrid_Regra_Col
            
                lErro = Saida_Celula_Regra(objGridInt)
                If lErro <> SUCESSO Then gError 185624
                
            Case iGrid_Percentual_Col
            
                lErro = Saida_Celula_Percentual(objGridInt)
                If lErro <> SUCESSO Then gError 185625
            
            Case iGrid_CondPagto_Col
            
                lErro = Saida_Celula_CondPagto(objGridInt)
                If lErro <> SUCESSO Then gError 185626
            
            Case iGrid_Observacao_Col
            
                lErro = Saida_Celula_Observacao(objGridInt)
                If lErro <> SUCESSO Then gError 185627

        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 185628
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 185624 To 185627
    
        Case 185628
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185629)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Regra(objGridInt As AdmGrid) As Long
'faz a critica da celula regra do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim colMnemonico As New Collection

On Error GoTo Erro_Saida_Celula_Regra

    Set objGridInt.objControle = Regra

    If Len(Trim(Regra.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
        
            lErro = CF("MnemonicoPRJ_Le", colMnemonico)
            If lErro <> SUCESSO Then gError 185630

            lErro = CF("Valida_Formula_WFW", Regra.Text, TIPO_DATA, iInicio, iTamanho, colMnemonico)
            If lErro <> SUCESSO Then gError 185631
                
        End If
        
        If GridRegras.Row - GridRegras.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185632

    Saida_Celula_Regra = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Regra:

    Saida_Celula_Regra = gErr
    
    Select Case gErr
    
        Case 185630, 185632
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 185631
            Regra.SelStart = iInicio
            Regra.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185633)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Percentual(objGridInt As AdmGrid) As Long
'faz a critica da celula Percentual do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = Percentual
                    
    'Se o campo foi preenchido
    If Len(Percentual.Text) > 0 Then

        'Critica o valor
        lErro = Porcentagem_Critica(Percentual.Text)
        If lErro <> SUCESSO Then gError 185693
            
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185678

    Saida_Celula_Percentual = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = gErr
    
    Select Case gErr
    
        Case 185678, 185693
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185679)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_CondPagto(objGridInt As AdmGrid) As Long
'faz a critica da celula CondPagto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_CondPagto

    Set objGridInt.objControle = CondPagto

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondPagto.Text)) <> 0 Then

        'Verifica se é uma Condicaopagamento selecionada
        If CondPagto.Text <> CondPagto.List(CondPagto.ListIndex) Then
    
            'Tenta selecionar na combo
            lErro = Combo_Seleciona_Grid(CondPagto, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 185695
            
            'Nao existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then
        
                objCondicaoPagto.iCodigo = iCodigo
        
                'Tenta ler CondicaoPagto com esse código no BD
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 185696
                
                'Não encontrou CondicaoPagto no BD
                If lErro <> SUCESSO Then gError 185697
        
                'Encontrou CondicaoPagto no BD e não é de Recebimento
                If objCondicaoPagto.iEmPagamento = 0 Then gError 185698
        
                'Coloca no Text da Combo
                CondPagto.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
        
            End If
        
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then gError 185699
            
            GridRegras.TextMatrix(GridRegras.Row, iGrid_CondPagto_Col) = CondPagto.Text
        
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185700

    Saida_Celula_CondPagto = SUCESSO

    Exit Function

Erro_Saida_Celula_CondPagto:

    Saida_Celula_CondPagto = gErr

    Select Case gErr
        
        Case 185695, 185696, 185700
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 185697
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")

            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)

            End If
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 185698
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 185699
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondPagto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185701)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'faz a critica da celula Observacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185682

    Saida_Celula_Observacao = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr
    
    Select Case gErr
    
        Case 185682
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185683)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 185634
    
    Call Limpa_Tela_RecebPRJ

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 185634
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185635)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
'grava os dados da tela

Dim lErro As Long
Dim objPRJReceb As New ClassPRJRecebPagto

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Projeto.ClipText)) = 0 Then gError 185702
    If Len(Trim(Numero.Text)) = 0 Then gError 185703
    If Len(Trim(Valor.Text)) = 0 Then gError 185704

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(objPRJReceb)
    If lErro <> SUCESSO Then gError 185705
    
    lErro = Critica_Dados(objPRJReceb)
    If lErro <> SUCESSO Then gError 185706

    lErro = Trata_Alteracao(objPRJReceb, objPRJReceb.lNumero, objPRJReceb.iFilialEmpresa, objPRJReceb.iTipo)
    If lErro <> SUCESSO Then gError 185707

    'Grava a etapa no Banco de Dados
    lErro = CF("PRJRecebPagto_Grava", objPRJReceb)
    If lErro <> SUCESSO Then gError 185708

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 185702
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Projeto.SetFocus

        Case 185703
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_PRJ_PAGTO_NAO_PREENCHIDO", gErr)
            Numero.SetFocus

        Case 185704
            Call Rotina_Erro(vbOKOnly, "ERRO_VLR_PRJ_PAGTO_NAO_PREENCHIDO", gErr)
            Valor.SetFocus
            
        Case 185705 To 185708

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185709)

    End Select

    Exit Function
    
End Function

Public Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objPRJReceb As New ClassPRJRecebPagto
    
On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Numero.Text)) = 0 Then gError 185714

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RECEBPRJ", Numero.Text)
    
    If vbMsgRes = vbYes Then
    
        objPRJReceb.iFilialEmpresa = giFilialEmpresa
        objPRJReceb.lNumero = StrParaLong(Numero.Text)
        objPRJReceb.iTipo = PRJ_TIPO_RECEB
    
        'exclui o modelo padrão de contabilização em questão
        lErro = CF("PRJRecebPagto_Exclui", objPRJReceb)
        If lErro <> SUCESSO Then gError 185637
    
        Call Limpa_Tela_RecebPRJ
        
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 185637
        
        Case 185714
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_PRJ_PAGTO_NAO_PREENCHIDO", gErr)
            Numero.SetFocus
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185638)
        
    End Select

    Exit Sub
    
End Sub

Function Limpa_Tela_RecebPRJ() As Long

    Call Grid_Limpa(objGridRegra)

    Call Limpa_Tela(Me)
    
    sProjetoAnt = ""
    sNomeProjetoAnt = ""
    
    Filial.Clear
    Etapa.Clear
    
    Mnemonicos.ListIndex = -1
    Funcoes.ListIndex = -1
    Operadores.ListIndex = -1
    
    Checkbox_Verifica_Sintaxe.Value = vbChecked
    CronFisFin.Value = vbUnchecked
    
    Limpa_Tela_RecebPRJ = SUCESSO
    
End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 185639

    Call Limpa_Tela_RecebPRJ
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 185639
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185640)
        
    End Select
    
End Sub

Public Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Mnemonicos_Click()

Dim iPos As Integer
Dim lErro As Long
Dim lPos As Long
Dim objMnemonico As New ClassMnemonicoPRJ
Dim sMnemonico As String

On Error GoTo Erro_Mnemonicos_Click
    
    If Len(Mnemonicos.Text) > 0 Then

        objMnemonico.sMnemonicoCombo = Mnemonicos.Text
    
        'retorna os dados do mnemonico passado como parametro
        lErro = CF("MnemonicoPRJ_Le_Mnemonico", objMnemonico)
        If lErro <> SUCESSO And lErro <> 178118 Then gError 185641

        If lErro = 178118 Then gError 185642
        
        Descricao.Text = objMnemonico.sMnemonicoDesc
        
        lPos = InStr(1, Mnemonicos.Text, "(")
        If lPos = 0 Then
            sMnemonico = Mnemonicos.Text
        Else
            sMnemonico = Mid(Mnemonicos.Text, 1, lPos)
        End If
        
        lErro = Mnemonicos1(sMnemonico)
        If lErro <> SUCESSO Then gError 185643
        
    End If
    
    Exit Sub
    
Erro_Mnemonicos_Click:

    Select Case gErr
    
        Case 185640, 185643
    
        Case 185642
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", gErr, objMnemonico.sMnemonicoCombo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185644)
            
    End Select
        
    Exit Sub
        
End Sub

Public Sub Operadores_Click()

Dim iPos As Integer
Dim lErro As Long
Dim objFormulaOperador As New ClassFormulaOperador
Dim lPos As Integer

On Error GoTo Erro_Operadores_Click
    
    objFormulaOperador.sOperadorCombo = Operadores.Text
    
    'retorna os dados do operador passado como parametro
    lErro = CF("FormulaOperador_Le", objFormulaOperador)
    If lErro <> SUCESSO And lErro <> 36098 Then gError 185645
    
    Descricao.Text = objFormulaOperador.sOperadorDesc
    
    Call Operadores1
    
    Exit Sub
    
Erro_Operadores_Click:

    Select Case gErr
    
        Case 185645
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185646)
            
    End Select
        
    Exit Sub

End Sub

Private Sub Posiciona_Texto_Tela(objControl As Control, sTexto As String)
'posiciona o texto sTexto no controle objControl da tela

Dim iPos As Integer
Dim iTamanho As Integer
Dim objGrid As Object

    iPos = objControl.SelStart
    objControl.Text = Mid(objControl.Text, 1, iPos) & sTexto & Mid(objControl.Text, iPos + 1, Len(objControl.Text))
    objControl.SelStart = iPos + Len(sTexto)
    
    If Not (Me.ActiveControl Is objControl) Then
    
        Set objGrid = GridRegras
    
        If iPos >= Len(objGrid.TextMatrix(objGrid.Row, objGrid.Col)) Then
            iTamanho = 0
        Else
            iTamanho = Len(objGrid.TextMatrix(objGrid.Row, objGrid.Col)) - iPos
        End If
        objGrid.TextMatrix(objGrid.Row, objGrid.Col) = Mid(objGrid.TextMatrix(objGrid.Row, objGrid.Col), 1, iPos) & sTexto & Mid(objGrid.TextMatrix(objGrid.Row, objGrid.Col), iPos + 1, iTamanho)
        
    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Funcoes1(sFuncao As String) As Long

Dim iPos As Integer

On Error GoTo Erro_Funcoes1

    If GridRegras.Row > 0 And GridRegras.Row <= objGridRegra.iLinhasExistentes + 1 And GridRegras.Col > 0 Then
        
        Select Case GridRegras.Col
        
            Case iGrid_Regra_Col
                Call Posiciona_Texto_Tela(Regra, Funcoes.Text)
                        
        End Select
        
    End If
    
    Funcoes1 = SUCESSO
    
    Exit Function
    
Erro_Funcoes1:

    Funcoes1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185647)
            
    End Select
        
    Exit Function

End Function

Function Mnemonicos1(sMnemonico As String) As Long

Dim iPos As Integer

On Error GoTo Erro_Mnemonicos1

    If GridRegras.Row > 0 And GridRegras.Row <= objGridRegra.iLinhasExistentes + 1 And GridRegras.Col > 0 Then
        
        Select Case GridRegras.Col
        
            Case iGrid_Regra_Col
            
                Call Posiciona_Texto_Tela(Regra, Mnemonicos.Text)
                
                If GridRegras.Row - GridRegras.FixedRows = objGridRegra.iLinhasExistentes Then
                    objGridRegra.iLinhasExistentes = objGridRegra.iLinhasExistentes + 1
                End If
                
                
        End Select
            
    End If

    Mnemonicos1 = SUCESSO
    
    Exit Function
    
Erro_Mnemonicos1:

    Mnemonicos1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185648)
            
    End Select
        
    Exit Function

End Function

Function Operadores1() As Long

Dim iPos As Integer

On Error GoTo Erro_Operadores1

    If GridRegras.Row > 0 And GridRegras.Row <= objGridRegra.iLinhasExistentes + 1 And GridRegras.Col > 0 Then
        
        Select Case GridRegras.Col
        
            Case iGrid_Regra_Col
                Call Posiciona_Texto_Tela(Regra, Operadores.Text)
                        
        End Select
        
    End If
     
    Operadores1 = SUCESSO
    
    Exit Function
    
Erro_Operadores1:

    Operadores1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185649)
            
    End Select
        
    Exit Function

End Function

Sub Projeto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim lNumIntDocPRJ As Long
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Projeto_Validate

    'Se alterou o projeto
    If sProjetoAnt <> Projeto.Text Then

        If Len(Trim(Projeto.ClipText)) > 0 Then
            
            lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError 189080
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185650
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185651
            
            lNumIntDocPRJ = objProjeto.lNumIntDoc
            
            NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
            
        End If
        
        sProjetoAnt = Projeto.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 185652
        
    End If
   
    Exit Sub

Erro_Projeto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 185650, 185652, 189080
        
        Case 185651
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185653)

    End Select

    Exit Sub

End Sub

Sub NomeReduzidoPrj_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim lNumIntDocPRJ As Long

On Error GoTo Erro_NomeReduzidoPrj_Validate

    'Se alterou o projeto
    If sNomeProjetoAnt <> NomeReduzidoPRJ.Text Then

        If Len(Trim(NomeReduzidoPRJ.Text)) > 0 Then
            
            objProjeto.sNomeReduzido = NomeReduzidoPRJ.Text
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le_NomeReduzido", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185654
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185655
            
            lNumIntDocPRJ = objProjeto.lNumIntDoc
            
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189117
            
        End If
        
        sNomeProjetoAnt = NomeReduzidoPRJ.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 185656
        
    End If
    
    Exit Sub

Erro_NomeReduzidoPrj_Validate:

    Cancel = True

    Select Case gErr
    
        Case 185654, 185656, 189117
        
        Case 185655
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO3", gErr, objProjeto.sNomeReduzido, objProjeto.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185657)

    End Select

    Exit Sub

End Sub

Sub LabelProjeto_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjeto_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Projeto.ClipText)) <> 0 Then

        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189081

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoPRJ, , "Código")

    Exit Sub

Erro_LabelProjeto_Click:

    Select Case gErr
    
        Case 189081

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185658)

    End Select

    Exit Sub
    
End Sub

Sub LabelNomeRedPRJ_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeRedPRJ_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(NomeReduzidoPRJ.Text)) <> 0 Then

        objProjeto.sNomeReduzido = NomeReduzidoPRJ.Text

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoPRJ, , "Nome Reduzido")

    Exit Sub

Erro_LabelNomeRedPRJ_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185659)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoPRJ_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoPRJ_evSelecao

    Set objProjeto = obj1

    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189118
    
    NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
    
    Call Projeto_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoPRJ_evSelecao:

    Select Case gErr
    
        Case 189118

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185660)

    End Select

    Exit Sub

End Sub

Private Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomeReduzidoPrj_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Function Trata_Projeto(ByVal lNumIntDocPRJ As Long) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos

On Error GoTo Erro_Trata_Projeto
    
    If lNumIntDocPRJ <> 0 Then

        objProjeto.lNumIntDoc = lNumIntDocPRJ
    
        lErro = CF("CarregaCombo_Etapas", objProjeto, Etapa)
        If lErro <> SUCESSO Then gError 185661
        
    End If
    
    sProjetoAnt = Projeto.Text
    sNomeProjetoAnt = NomeReduzidoPRJ.Text
    
    Proposta.Text = ""

    Trata_Projeto = SUCESSO

    Exit Function

Erro_Trata_Projeto:

    Trata_Projeto = gErr

    Select Case gErr
    
        Case 185661

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185662)

    End Select

    Exit Function

End Function

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_ClienteLabel_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(Cliente.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objCliente.lCodigo = LCodigo_Extrai(Cliente.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objCliente.sNomeReduzido = Cliente.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente, "", sOrdenacao)

    Exit Sub
    
Erro_ClienteLabel_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155003)
    
    End Select
    
End Sub

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    
    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 129422

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 129423

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
       
    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If
        
    Exit Sub

Erro_Cliente_Validate:
        
    Cancel = True

    Select Case gErr
    
        Case 129422, 129423
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155004)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer
Dim sNomeRed As String
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Filial_Validate
        
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Filial
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129418

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Cliente
        If Len(Trim(Cliente.Text)) = 0 Then gError 129419

        'Lê o Cliente que está na tela
        sNomeRed = Trim(Cliente.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialCliente.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Cliente e Código da Filial
        lErro = CF("Filial_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 129420

        'Se não existe a Filial
        If lErro = 17660 Then gError 129421

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 129505
    
    Exit Sub
    
Erro_Filial_Validate:

    Select Case gErr

        Case 129418, 129420

        Case 129419
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 129421
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE1", Filial.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 129505
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155005)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call Cliente_Validate(bCancel)

    Exit Sub

End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Filial_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Valor_Validate
    
    'Verifica se algum valor foi digitado
    If Len(Trim(Valor.ClipText)) <> 0 Then

        'Critica se é valor positivo
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 185674
    
        'Põe o valor formatado na tela
        Valor.Text = Format(Valor.Text, "Standard")
    
    End If
    
    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case Err

        Case 185674
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185675)

    End Select

    Exit Sub

End Sub

Private Sub LabelNumero_Click()

Dim objPRJReceb As New ClassPRJRecebPagto
Dim colSelecao As New Collection

    objPRJReceb.lNumero = StrParaLong(Numero.Text)

    Call Chama_Tela("RecebimentoPRJLista", colSelecao, objPRJReceb, objEventoNumero)

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPRJReceb As ClassPRJRecebPagto

On Error GoTo Erro_objEventoNumero_evSelecao:

    Set objPRJReceb = obj1
    
    objPRJReceb.iFilialEmpresa = giFilialEmpresa
    objPRJReceb.iTipo = PRJ_TIPO_RECEB

    lErro = Traz_Recebimento_Tela(objPRJReceb)
    If lErro <> SUCESSO Then gError 185676

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr
    
        Case 185676
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185677)

    End Select

    Exit Sub
End Sub

Private Function Carrega_CondicaoPagamento(objCombo As ComboBox) As Long

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 185689

   For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na List da Combo CondicaoPagamento
        objCombo.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        objCombo.ItemData(objCombo.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case 185689

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185690)

    End Select

    Exit Function

End Function

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("PRJRecebPagto_Automatico", lCodigo, PRJ_TIPO_RECEB)
    If lErro <> SUCESSO Then gError 185724
    
    Numero.PromptInclude = False
    Numero.Text = CStr(lCodigo)
    Numero.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 185724

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185725)
    
    End Select

    Exit Sub
    
End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 185736

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 185736

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185737)

    End Select
    
    Exit Sub

End Sub

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objPRJPagto As New ClassPRJRecebPagto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RecebimentoPRJ"

    'Lê os dados da Tela PedidoVenda
    objPRJPagto.lNumero = StrParaLong(Numero.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Numero", objPRJPagto.lNumero, 0, "Numero"
    'Filtros para o Sistema de Setas
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185732)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objPRJPagto As New ClassPRJRecebPagto

On Error GoTo Erro_Tela_Preenche

    objPRJPagto.lNumero = colCampoValor.Item("Numero").vValor
    objPRJPagto.iFilialEmpresa = giFilialEmpresa
    objPRJPagto.iTipo = PRJ_TIPO_RECEB

    If objPRJPagto.lNumero <> 0 Then
        lErro = Traz_Recebimento_Tela(objPRJPagto)
        If lErro <> SUCESSO Then gError 185733
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 185733

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185734)

    End Select

    Exit Function

End Function

Private Sub Etapa_Click()

Dim lErro As Long
Dim sEtapa As String
    
On Error GoTo Erro_Funcoes_Click
    
    sEtapa = """ & SCodigo_Extrai(Etapa.Text) & """
    
    lErro = Funcoes1(sEtapa)
    If lErro <> SUCESSO Then gError 185609
    
    Exit Sub
    
Erro_Funcoes_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185610)
            
    End Select
        
    Exit Sub

End Sub

Private Sub LabelProposta_Click()

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas
Dim colSelecao As New Collection
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProposta_Click

    If Len(Trim(Projeto.ClipText)) = 0 Then gError 187546

    lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189082

    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa
    
    'Le
    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187547
    
    'Se não encontrou => Erro
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187548
    
    colSelecao.Add objProjeto.lNumIntDoc

    'Verifica se o Proposta foi preenchido
    If Len(Trim(Proposta.Text)) <> 0 Then

        objProposta.sCodigo = Proposta.Text

    End If

    Call Chama_Tela("PRJPropostasLista", colSelecao, objProposta, objEventoProposta, "NumIntDocPRJ = ?", "Código")

    Exit Sub

Erro_LabelProposta_Click:

    Select Case gErr
    
        Case 187546
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Projeto.SetFocus
            
        Case 187547, 189082

        Case 187548
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187549)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProposta_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas

On Error GoTo Erro_objEventoProposta_evSelecao

    Set objProposta = obj1
    
    Proposta.Text = objProposta.sCodigo
    Call Proposta_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoProposta_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187550)

    End Select

    Exit Sub

End Sub

Private Sub Proposta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Proposta_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProposta As New ClassPRJPropostas
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Proposta_Validate

    If Len(Trim(Proposta.Text)) > 0 Then
    
        If Len(Trim(Projeto.ClipText)) = 0 Then gError 187551
    
        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189083
    
        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
        'Le
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187552
        
        'Se não encontrou => Erro
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187553
        
        objProposta.lNumIntDocPRJ = objProjeto.lNumIntDoc
        objProposta.sCodigo = Proposta.Text
        
        'Lê a proposta que está sendo Passado
        lErro = CF("PRJPropostas_Le", objProposta)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187554
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187555
        
        Valor.Text = Format(objProposta.dValorTotal, "STANDARD")
        
        If objProposta.lCliente <> 0 Then
        
            Cliente.Text = objProposta.lCliente
            Call Cliente_Validate(bSGECancelDummy)
            
            If objProposta.iFilialCliente <> 0 Then
            
                Filial.Text = objProposta.iFilialCliente
                Call Filial_Validate(bSGECancelDummy)
                
            End If
            
        End If
        
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189120
        
        Projeto_Validate (bSGECancelDummy)
    
    End If
    
    Exit Sub

Erro_Proposta_Validate:

    Cancel = True

    Select Case gErr

        Case 187551
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Projeto.SetFocus
            
        Case 187552, 187554, 189083, 189120

        Case 187553
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case 187555
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJPROPOSTAS_NAO_CADASTRADO", gErr, Proposta.Text, objProjeto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187556)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelContrato_Click()

Dim lErro As Long
Dim objContrato As New ClassPRJContratos
Dim colSelecao As New Collection
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelContrato_Click

    If Len(Trim(Projeto.ClipText)) = 0 Then gError 187546

    lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189082

    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa
    
    'Le
    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187547
    
    'Se não encontrou => Erro
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187548
    
    colSelecao.Add objProjeto.lNumIntDoc

    'Verifica se o Contrato foi preenchido
    If Len(Trim(Contrato.Text)) <> 0 Then

        objContrato.sCodigo = Contrato.Text

    End If

    Call Chama_Tela("PRJContratosLista", colSelecao, objContrato, objEventoContrato, "NumIntDocPRJ = ?", "Código")

    Exit Sub

Erro_LabelContrato_Click:

    Select Case gErr
    
        Case 187546
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Projeto.SetFocus
            
        Case 187547, 189082

        Case 187548
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187549)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContrato_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objContrato As New ClassPRJContratos

On Error GoTo Erro_objEventoContrato_evSelecao

    Set objContrato = obj1
    
    Contrato.Text = objContrato.sCodigo
    Call Contrato_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoContrato_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187550)

    End Select

    Exit Sub

End Sub

Private Sub Contrato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Contrato_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContrato As New ClassPRJContratos
Dim objProjeto As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer
Dim objProposta As New ClassPRJPropostas

On Error GoTo Erro_Contrato_Validate

    If Len(Trim(Contrato.Text)) > 0 Then
    
        If Len(Trim(Projeto.ClipText)) = 0 Then gError 187551
    
        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189083
    
        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa
        
        'Le
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187552
        
        'Se não encontrou => Erro
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187553
        
        objContrato.lNumIntDocPRJ = objProjeto.lNumIntDoc
        objContrato.sCodigo = Contrato.Text
        
        'Lê a Contrato que está sendo Passado
        lErro = CF("PRJContratos_Le", objContrato)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187554
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 187555
        
        objProposta.lNumIntDoc = objContrato.lNumIntDocProposta
        
        lErro = CF("PRJPropostas_Le_NumIntDoc", objProposta)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 189435
        
        Proposta.Text = objProposta.sCodigo
        
        Valor.Text = Format(objContrato.dValorTotal, "STANDARD")
        
        If objContrato.lCliente <> 0 Then
        
            Cliente.Text = objContrato.lCliente
            Call Cliente_Validate(bSGECancelDummy)
            
            If objContrato.iFilialCliente <> 0 Then
            
                Filial.Text = objContrato.iFilialCliente
                Call Filial_Validate(bSGECancelDummy)
                
            End If
            
        End If
                
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189120
        
        Projeto_Validate (bSGECancelDummy)
    
    End If
    
    Exit Sub

Erro_Contrato_Validate:

    Cancel = True

    Select Case gErr

        Case 187551
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Projeto.SetFocus
            
        Case 187552, 187554, 189083, 189120, 189435

        Case 187553
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case 187555
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJContratoS_NAO_CADASTRADO", gErr, Contrato.Text, objProjeto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187556)

    End Select

    Exit Sub
    
End Sub
