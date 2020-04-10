VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PagamentoPRJ 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9315
   Begin VB.CheckBox CronFisFin 
      Caption         =   "Inclui no Cronograma Fís. Financ."
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
      Left            =   4680
      TabIndex        =   8
      Top             =   1290
      Value           =   1  'Checked
      Width           =   3225
   End
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
      Left            =   300
      TabIndex        =   7
      Top             =   1320
      Value           =   1  'Checked
      Width           =   3285
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4695
      TabIndex        =   6
      Top             =   915
      Width           =   1815
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2040
      Picture         =   "PagamentoPRJ.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Numeração Automática"
      Top             =   570
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condições"
      Height          =   4365
      Left            =   300
      TabIndex        =   19
      Top             =   1575
      Width           =   8850
      Begin VB.TextBox Observacao 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   2175
         Width           =   2505
      End
      Begin VB.ComboBox Etapa 
         Height          =   315
         ItemData        =   "PagamentoPRJ.ctx":00EA
         Left            =   2895
         List            =   "PagamentoPRJ.ctx":00EC
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3240
         Width           =   2025
      End
      Begin VB.TextBox Regra 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3210
         TabIndex        =   23
         Top             =   705
         Width           =   3240
      End
      Begin VB.ComboBox CondPagto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4665
         TabIndex        =   22
         Top             =   1605
         Width           =   1455
      End
      Begin VB.ComboBox Operadores 
         Height          =   315
         Left            =   7620
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3270
         Width           =   1050
      End
      Begin VB.ComboBox Funcoes 
         Height          =   315
         Left            =   5055
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3255
         Width           =   2415
      End
      Begin VB.ComboBox Mnemonicos 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PagamentoPRJ.ctx":00EE
         Left            =   195
         List            =   "PagamentoPRJ.ctx":00FB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3240
         Width           =   2520
      End
      Begin VB.TextBox Descricao 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   540
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   3675
         Width           =   8475
      End
      Begin MSMask.MaskEdBox Percentual 
         Height          =   270
         Left            =   4650
         TabIndex        =   21
         Top             =   1140
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
      Begin MSFlexGridLib.MSFlexGrid GridRegras 
         Height          =   2625
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   4630
         _Version        =   393216
         Rows            =   50
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
         Left            =   7650
         TabIndex        =   27
         Top             =   3000
         Width           =   1050
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
         Left            =   195
         TabIndex        =   26
         Top             =   3000
         Width           =   1125
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
         TabIndex        =   25
         Top             =   3015
         Width           =   795
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
         Left            =   2895
         TabIndex        =   24
         Top             =   3000
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7005
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   18
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PagamentoPRJ.ctx":011D
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "PagamentoPRJ.ctx":029B
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PagamentoPRJ.ctx":07CD
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PagamentoPRJ.ctx":0957
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   300
      Left            =   1260
      TabIndex        =   5
      Top             =   915
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   300
      Left            =   1260
      TabIndex        =   2
      Top             =   555
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
      Left            =   4695
      TabIndex        =   4
      Top             =   510
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
      Left            =   1260
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
      Left            =   4695
      TabIndex        =   1
      Top             =   120
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
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
      Left            =   3420
      TabIndex        =   33
      Top             =   165
      Width           =   1410
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
      Left            =   555
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   32
      Top             =   165
      Width           =   675
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4110
      TabIndex        =   31
      Top             =   960
      Width           =   525
   End
   Begin VB.Label FornLabel 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
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
      Left            =   195
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   30
      Top             =   960
      Width           =   1035
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
      Left            =   510
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   29
      Top             =   585
      Width           =   720
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
      Left            =   4125
      TabIndex        =   28
      Top             =   540
      Width           =   510
   End
End
Attribute VB_Name = "PagamentoPRJ"
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
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

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
    Caption = "Pagamentos relacionados a Projeto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagamentoPRJ"
    
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
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornLabel_Click
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
    Set objEventoFornecedor = Nothing
    
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
    Set objEventoFornecedor = New AdmEvento
    
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
    If lErro <> SUCESSO Then gError 189064
       
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 185612 To 185614, 185691, 189064
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185615)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional ByVal objPRJPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objPRJPagto Is Nothing) Then
    
        objPRJPagto.iFilialEmpresa = giFilialEmpresa
        objPRJPagto.iTipo = PRJ_TIPO_PAGTO
    
        lErro = Traz_Pagamento_Tela(objPRJPagto)
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

Function Traz_Pagamento_Tela(ByVal objPRJPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim objPRJPagtoRegras As ClassPRJRecebPagtoRegras
Dim iLinha As Integer

On Error GoTo Erro_Traz_Pagamento_Tela

    Call Limpa_Tela_PagtoPRJ

    'Lê a Etapa que está sendo Passada
    lErro = CF("PRJRecebPagto_Le", objPRJPagto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185687

    If objPRJPagto.lNumero <> 0 Then
        Numero.PromptInclude = False
        Numero.Text = objPRJPagto.lNumero
        Numero.PromptInclude = True
    End If
    If objPRJPagto.dValor <> 0 Then Valor.Text = Format(objPRJPagto.dValor, "STANDARD")
    
    If objPRJPagto.lCliForn <> 0 Then
        Fornecedor.Text = objPRJPagto.lCliForn
        Call Fornecedor_Validate(bSGECancelDummy)
    End If
    
    Call Combo_Seleciona(Filial, objPRJPagto.iFilial)

    If lErro = SUCESSO Then
    
        objProjeto.lNumIntDoc = objPRJPagto.lNumIntDocPRJ
        
        'Lê o Projetos que está sendo Passado
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185688
        
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189121
        
        Call Projeto_Validate(bSGECancelDummy)
        
        If objPRJPagto.iIncluiCFF = MARCADO Then
            CronFisFin.Value = vbChecked
        Else
            CronFisFin.Value = vbUnchecked
        End If
        
        iLinha = 0
        For Each objPRJPagtoRegras In objPRJPagto.colRegras
        
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
    
        objGridRegra.iLinhasExistentes = objPRJPagto.colRegras.Count
        
    Else
    
        If objPRJPagto.lNumIntDocPRJ <> 0 Then
    
            objProjeto.lNumIntDoc = objPRJPagto.lNumIntDocPRJ
            
            'Lê o Projetos que está sendo Passado
            lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185688
            
            lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
            If lErro <> SUCESSO Then gError 189128
            
            Call Projeto_Validate(bSGECancelDummy)
            
        End If

    End If
    
    Traz_Pagamento_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Pagamento_Tela:

    Traz_Pagamento_Tela = gErr

    Select Case gErr
    
        Case 185687 To 185688, 185692, 189121, 189128
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185616)
    
    End Select
    
    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objPRJPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim objPRJPagtoRegra As ClassPRJRecebPagtoRegras
Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer
Dim objProjeto As ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    Set objProjeto = New ClassProjetos

    lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189084

    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa
    
    'Le
    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185728
    
    'Se não encontrou => Erro
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185729
    
    objPRJPagto.lNumIntDocPRJ = objProjeto.lNumIntDoc

    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.ClipText)) > 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê o Fornecedor através do Nome Reduzido
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 129491

        If lErro = SUCESSO Then objPRJPagto.lCliForn = objFornecedor.lCodigo
                            
    End If
    
    objPRJPagto.iTipo = PRJ_TIPO_PAGTO
    objPRJPagto.iFilialEmpresa = giFilialEmpresa
    objPRJPagto.iFilial = Codigo_Extrai(Filial.Text)
    objPRJPagto.dValor = StrParaDbl(Valor.Text)
    objPRJPagto.lNumero = StrParaLong(Numero.Text)

    If CronFisFin.Value = vbChecked Then
        objPRJPagto.iIncluiCFF = MARCADO
    Else
        objPRJPagto.iIncluiCFF = DESMARCADO
    End If
    
    For iIndice = 1 To objGridRegra.iLinhasExistentes
    
        Set objPRJPagtoRegra = New ClassPRJRecebPagtoRegras
        
        objPRJPagtoRegra.sObservacao = GridRegras.TextMatrix(iIndice, iGrid_Observacao_Col)
        objPRJPagtoRegra.sRegra = GridRegras.TextMatrix(iIndice, iGrid_Regra_Col)
        objPRJPagtoRegra.iCondPagto = Codigo_Extrai(GridRegras.TextMatrix(iIndice, iGrid_CondPagto_Col))
        objPRJPagtoRegra.dPercentual = StrParaDbl(Val(GridRegras.TextMatrix(iIndice, iGrid_Percentual_Col)) / 100)
        
        objPRJPagto.colRegras.Add objPRJPagtoRegra
    
    Next
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 185728, 189084
    
        Case 185729
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185684)
    
    End Select
    
    Exit Function

End Function

Function Critica_Dados(ByVal objPRJPagto As ClassPRJRecebPagto) As Long

Dim lErro As Long
Dim objPRJPagtoRegra As ClassPRJRecebPagtoRegras
Dim iLinha As Integer
Dim dPercent As Double
Dim colMnemonico As New Collection
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Critica_Dados

    lErro = CF("MnemonicoPRJ_Le", colMnemonico)
    If lErro <> SUCESSO Then gError 189372
    
    iLinha = 0
    For Each objPRJPagtoRegra In objPRJPagto.colRegras
    
        iLinha = iLinha + 1
        If Len(Trim(objPRJPagtoRegra.sRegra)) = 0 Then gError 185710
        If objPRJPagtoRegra.iCondPagto = 0 Then gError 185711
        If objPRJPagtoRegra.dPercentual = 0 Then gError 185712
    
        dPercent = dPercent + objPRJPagtoRegra.dPercentual
    
        lErro = CF("Valida_Formula_WFW", objPRJPagtoRegra.sRegra, TIPO_DATA, iInicio, iTamanho, colMnemonico)
        If lErro <> SUCESSO Then gError 189373
        
    Next
    
    If Abs(dPercent - 1) > QTDE_ESTOQUE_DELTA2 Then gError 189016
    
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
    
        Case 189016
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENT_TOTAL_NAO_100PERC", gErr)
    
        Case 189372, 189373
    
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
    
    Call Limpa_Tela_PagtoPRJ

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
Dim objPRJPagto As New ClassPRJRecebPagto

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Projeto.ClipText)) = 0 Then gError 185702
    If Len(Trim(Numero.Text)) = 0 Then gError 185703
    If Len(Trim(Valor.Text)) = 0 Then gError 185704

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(objPRJPagto)
    If lErro <> SUCESSO Then gError 185705
    
    lErro = Critica_Dados(objPRJPagto)
    If lErro <> SUCESSO Then gError 185706

    lErro = Trata_Alteracao(objPRJPagto, objPRJPagto.lNumero, objPRJPagto.iFilialEmpresa, objPRJPagto.iTipo)
    If lErro <> SUCESSO Then gError 185707

    'Grava a etapa no Banco de Dados
    lErro = CF("PRJRecebPagto_Grava", objPRJPagto)
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
Dim objPagtoPRJ As New ClassPRJRecebPagto
    
On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Numero.Text)) = 0 Then gError 185714

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_PAGTOPRJ", Numero.Text)
    
    If vbMsgRes = vbYes Then
    
        objPagtoPRJ.iFilialEmpresa = giFilialEmpresa
        objPagtoPRJ.lNumero = StrParaLong(Numero.Text)
        objPagtoPRJ.iTipo = PRJ_TIPO_PAGTO
    
        'exclui o modelo padrão de contabilização em questão
        lErro = CF("PRJRecebPagto_Exclui", objPagtoPRJ)
        If lErro <> SUCESSO Then gError 185637
    
        Call Limpa_Tela_PagtoPRJ
        
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

Function Limpa_Tela_PagtoPRJ() As Long

    Call Grid_Limpa(objGridRegra)

    Call Limpa_Tela(Me)
    
    sProjetoAnt = ""
    sNomeProjetoAnt = ""
    
    Checkbox_Verifica_Sintaxe.Value = vbChecked
    CronFisFin.Value = vbUnchecked
    
    Filial.Clear
    Etapa.Clear
    
    Mnemonicos.ListIndex = -1
    Funcoes.ListIndex = -1
    Operadores.ListIndex = -1
    
    Limpa_Tela_PagtoPRJ = SUCESSO
    
End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 185639

    Call Limpa_Tela_PagtoPRJ
    
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
                Call Posiciona_Texto_Tela(Regra, sFuncao)
                        
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
            If lErro <> SUCESSO Then gError 189085
            
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
    
        Case 185650, 185652, 189085
        
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
            If lErro <> SUCESSO Then gError 189122
            
        End If
        
        sNomeProjetoAnt = NomeReduzidoPRJ.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 185656
        
    End If
    
    Exit Sub

Erro_NomeReduzidoPrj_Validate:

    Cancel = True

    Select Case gErr
    
        Case 185654, 185656, 189122
        
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
        If lErro <> SUCESSO Then gError 189086

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoPRJ, , "Código")

    Exit Sub

Erro_LabelProjeto_Click:

    Select Case gErr
    
        Case 189086

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
    If lErro <> SUCESSO Then gError 189123
    
    NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
    
    Call Projeto_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoPRJ_evSelecao:

    Select Case gErr
    
        Case 189123

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

Private Sub FornLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    'Preenche objFornecedor com NomeReduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    
    Call Fornecedor_Preenche

End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iCondPagto As Integer
Dim objTipoFornecedor As New ClassTipoFornecedor

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se fornecedor esta preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 185663

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6698 Then gError 185664

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
    Else

        Filial.Clear

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 185663, 185664

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 185665)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Dispara Validate de Fornecedor
    Call Fornecedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Se a filial nao estiver preenchida => sai da rotina
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 185666

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 185667

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 185668

        If lErro = 18272 Then gError 185669

        'coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 185670

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 185666, 185668

        Case 185667
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 185669
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 185670
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185671)

    End Select

    Exit Sub

End Sub

Public Sub Fornecedor_Preenche()

Static sNomeReduzidoParte As String
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 185672

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 185672

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185673)

    End Select
    
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

Dim objPRJPagto As New ClassPRJRecebPagto
Dim colSelecao As New Collection

    'Preenche objFornecedor com NomeReduzido da tela
    objPRJPagto.lNumero = StrParaLong(Numero.Text)

    Call Chama_Tela("PagamentoPRJLista", colSelecao, objPRJPagto, objEventoNumero)

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPRJPagto As ClassPRJRecebPagto

On Error GoTo Erro_objEventoNumero_evSelecao:

    Set objPRJPagto = obj1
    
    objPRJPagto.iFilialEmpresa = giFilialEmpresa
    objPRJPagto.iTipo = PRJ_TIPO_PAGTO

    lErro = Traz_Pagamento_Tela(objPRJPagto)
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
    
    lErro = CF("PRJRecebPagto_Automatico", lCodigo, PRJ_TIPO_PAGTO)
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

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objPRJPagto As New ClassPRJRecebPagto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PagamentoPRJ"

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
    objPRJPagto.iTipo = PRJ_TIPO_PAGTO

    If objPRJPagto.lNumero <> 0 Then
        lErro = Traz_Pagamento_Tela(objPRJPagto)
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
    
    sEtapa = """" & SCodigo_Extrai(Etapa.Text) & """"
    
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

