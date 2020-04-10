VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl CobrancaEmailPadraoOcx 
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   9315
   Begin VB.Frame FrameAtraso 
      Caption         =   "Válido para atrasos"
      Height          =   1725
      Left            =   7035
      TabIndex        =   39
      Top             =   1215
      Width           =   2130
      Begin MSMask.MaskEdBox AtrasoDe 
         Height          =   315
         Left            =   615
         TabIndex        =   40
         Top             =   405
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         AllowPrompt     =   -1  'True
         MaxLength       =   4
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AtrasoAte 
         Height          =   315
         Left            =   630
         TabIndex        =   41
         Top             =   960
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         AllowPrompt     =   -1  'True
         MaxLength       =   4
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   45
         Top             =   450
         Width           =   315
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   44
         Top             =   1005
         Width           =   360
      End
      Begin VB.Label Label7 
         Caption         =   "dias"
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
         Height          =   270
         Left            =   1410
         TabIndex        =   43
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "dias"
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
         Height          =   270
         Left            =   1410
         TabIndex        =   42
         Top             =   1020
         Width           =   390
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Outros"
      Height          =   1725
      Left            =   195
      TabIndex        =   33
      Top             =   1215
      Width           =   6765
      Begin VB.ComboBox Usuario 
         Height          =   315
         Left            =   1590
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Modelo 
         Height          =   315
         Left            =   1590
         MaxLength       =   250
         TabIndex        =   4
         Top             =   210
         Width           =   4515
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6120
         TabIndex        =   34
         Top             =   210
         Width           =   495
      End
      Begin MSMask.MaskEdBox De 
         Height          =   315
         Left            =   1590
         TabIndex        =   5
         Top             =   570
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeExibicao 
         Height          =   315
         Left            =   1590
         TabIndex        =   6
         Top             =   945
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Uso Exclusivo:"
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
         Index           =   19
         Left            =   150
         TabIndex        =   46
         Top             =   1365
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome Exibição:"
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
         Index           =   13
         Left            =   90
         TabIndex        =   37
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Index           =   12
         Left            =   1110
         TabIndex        =   36
         Top             =   615
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "Modelo Html:"
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
         Index           =   3
         Left            =   300
         TabIndex        =   35
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificação"
      Height          =   1080
      Left            =   195
      TabIndex        =   29
      Top             =   75
      Width           =   6750
      Begin VB.ComboBox Tipo 
         Height          =   315
         ItemData        =   "CobrancaEmailPadrao.ctx":0000
         Left            =   3345
         List            =   "CobrancaEmailPadrao.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   195
         Width           =   3315
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2160
         Picture         =   "CobrancaEmailPadrao.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   225
         Width           =   300
      End
      Begin MSMask.MaskEdBox DescReduzida 
         Height          =   315
         Left            =   1605
         TabIndex        =   3
         Top             =   615
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1605
         TabIndex        =   0
         Top             =   210
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Index           =   10
         Left            =   2865
         TabIndex        =   38
         Top             =   240
         Width           =   450
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
         Left            =   780
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   31
         Top             =   255
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Index           =   2
         Left            =   510
         TabIndex        =   30
         Top             =   660
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Email Padrão"
      Height          =   3885
      Left            =   195
      TabIndex        =   21
      Top             =   2955
      Width           =   8985
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
         Left            =   1560
         TabIndex        =   7
         Top             =   180
         Value           =   1  'Checked
         Width           =   3285
      End
      Begin VB.TextBox Anexo 
         Height          =   315
         Left            =   1575
         MaxLength       =   250
         TabIndex        =   11
         Top             =   1185
         Width           =   7155
      End
      Begin VB.TextBox Mensagem 
         Height          =   1080
         Left            =   1575
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1560
         Width           =   7155
      End
      Begin VB.TextBox Assunto 
         Height          =   315
         Left            =   1575
         MaxLength       =   250
         TabIndex        =   10
         Top             =   810
         Width           =   7155
      End
      Begin VB.TextBox Cc 
         Height          =   315
         Left            =   1575
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   450
         Width           =   7155
      End
      Begin VB.ComboBox Operadores 
         Height          =   315
         Left            =   7275
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2895
         Width           =   1515
      End
      Begin VB.ComboBox Funcoes 
         Height          =   315
         Left            =   4185
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2880
         Width           =   2925
      End
      Begin VB.ComboBox Mnemonicos 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "CobrancaEmailPadrao.ctx":00EE
         Left            =   195
         List            =   "CobrancaEmailPadrao.ctx":00FB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2865
         Width           =   3780
      End
      Begin VB.TextBox Descricao 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   585
         Left            =   195
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   3225
         Width           =   8580
      End
      Begin VB.Label Label1 
         Caption         =   "Anexo:"
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
         Index           =   11
         Left            =   900
         TabIndex        =   32
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Mensagem CRM:"
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
         Index           =   6
         Left            =   60
         TabIndex        =   28
         Top             =   1605
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Assunto:"
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
         Index           =   5
         Left            =   765
         TabIndex        =   27
         Top             =   825
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Cc:"
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
         Index           =   4
         Left            =   1230
         TabIndex        =   26
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
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
         Index           =   9
         Left            =   7305
         TabIndex        =   25
         Top             =   2670
         Width           =   1050
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   195
         TabIndex        =   24
         Top             =   2670
         Width           =   1125
      End
      Begin VB.Label Label1 
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
         Index           =   8
         Left            =   4185
         TabIndex        =   23
         Top             =   2685
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7005
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   20
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CobrancaEmailPadrao.ctx":011D
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CobrancaEmailPadrao.ctx":029B
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CobrancaEmailPadrao.ctx":07CD
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CobrancaEmailPadrao.ctx":0957
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "CobrancaEmailPadraoOcx"
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

Dim iAlterado As Integer
Dim iFocus As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Const FOCUS_CC = 1
Const FOCUS_ASSUNTO = 2
Const FOCUS_MENSAGEM = 3
Const FOCUS_ANEXO = 4

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
    Caption = "Modelo de Emails"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = "CobrancaEmailPadrao"
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

Private Sub De_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Modelo_Change()
'
'Dim vbResult As VbMsgBoxResult
'
'    If Len(Trim(Modelo.Text)) > 0 Then
'        If Len(Trim(Mensagem.Text)) > 0 Then
'            vbResult = Rotina_Aviso(vbYesNo, "AVISO_MODELO_COM_MENSAGEM")
'            If vbResult = vbNo Then
'                Modelo.Text = ""
'            Else
'                Mensagem.Text = ""
'                Mensagem.Enabled = False
'            End If
'        End If
'    Else
'        Mensagem.Enabled = True
'    End If

End Sub

Private Sub NomeExibicao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

'***** fim do trecho a ser copiado ******
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objEventoCodigo = Nothing
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
    If lErro <> SUCESSO And lErro <> 36088 Then gError 187053
    
    Descricao.Text = objFormulaFuncao.sFuncaoDesc
    
    lPos = InStr(1, Funcoes.Text, "(")
    If lPos = 0 Then
        sFuncao = Funcoes.Text
    Else
        sFuncao = Mid(Funcoes.Text, 1, lPos)
    End If
    
    lErro = Funcoes1(sFuncao)
    If lErro <> SUCESSO Then gError 187054
    
    Exit Sub
    
Erro_Funcoes_Click:

    Select Case gErr
    
        Case 187053, 187054
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187055)
            
    End Select
        
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    
    'carrega a combobox de funcoes
    lErro = Carga_Combobox_Funcoes()
    If lErro <> SUCESSO Then gError 187056
    
    'carrega a combobox de operadores
    lErro = Carga_Combobox_Operadores()
    If lErro <> SUCESSO Then gError 187057
    
    'Carrega os mnemônicos de projetos
    lErro = Carga_Combobox_Mnemonicos
    If lErro <> SUCESSO Then gError 187058
    
    lErro = Carrega_Usuarios
    If lErro <> SUCESSO Then gError 187058
    
    Tipo.Clear
      
    Tipo.AddItem TIPO_COBRANCAEMAILPADRAO_COBRANCA & SEPARADOR & STRING_TIPO_COBRANCAEMAILPADRAO_COBRANCA
    Tipo.ItemData(Tipo.NewIndex) = TIPO_COBRANCAEMAILPADRAO_COBRANCA
          
    Tipo.AddItem TIPO_COBRANCAEMAILPADRAO_AVISO & SEPARADOR & STRING_TIPO_COBRANCAEMAILPADRAO_AVISO
    Tipo.ItemData(Tipo.NewIndex) = TIPO_COBRANCAEMAILPADRAO_AVISO
          
    Tipo.AddItem TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO & SEPARADOR & STRING_TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO
    Tipo.ItemData(Tipo.NewIndex) = TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO
    
    Tipo.AddItem TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP & SEPARADOR & STRING_TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP
    Tipo.ItemData(Tipo.NewIndex) = TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP
    
    Tipo.AddItem TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA & SEPARADOR & STRING_TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA
    Tipo.ItemData(Tipo.NewIndex) = TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA
    
    Tipo.AddItem TIPO_COBRANCAEMAILPADRAO_CONTATO_CLIENTE & SEPARADOR & STRING_TIPO_COBRANCAEMAILPADRAO_CONTATO_CLIENTE
    Tipo.ItemData(Tipo.NewIndex) = TIPO_COBRANCAEMAILPADRAO_CONTATO_CLIENTE
    
    Call Combo_Seleciona_ItemData(Tipo, TIPO_COBRANCAEMAILPADRAO_AVISO)
          
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 187056, 187057, 187058
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187059)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional ByVal objEmail As ClassCobrancaEmailPadrao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objEmail Is Nothing) Then
    
        lErro = Traz_Email_Tela(objEmail)
        If lErro <> SUCESSO Then gError 187060
    
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 187060
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187061)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Function Traz_Email_Tela(ByVal objEmail As ClassCobrancaEmailPadrao) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Email_Tela

    lErro = CF("CobrancaEmailPadrao_Le", objEmail)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187062
    
    Codigo.PromptInclude = False
    Codigo.Text = objEmail.lCodigo
    Codigo.PromptInclude = True
    
    DescReduzida.Text = objEmail.sDescricao

    If objEmail.iTipo = TIPO_COBRANCAEMAILPADRAO_COBRANCA Or objEmail.iTipo = TIPO_COBRANCAEMAILPADRAO_AVISO Then
        FrameAtraso.Visible = True
        FrameAtraso.Enabled = True
        AtrasoDe.Text = CStr(objEmail.iAtrasoDe)
        AtrasoAte.Text = CStr(objEmail.iAtrasoAte)
    Else
        FrameAtraso.Visible = False
        FrameAtraso.Enabled = False
        AtrasoDe.Text = ""
        AtrasoAte.Text = ""
    End If
    
    Cc.Text = objEmail.sCC
    Assunto.Text = objEmail.sAssunto
    Anexo.Text = objEmail.sAnexo
    Mensagem.Text = objEmail.sMensagem
    
    Call Combo_Seleciona_ItemData(Tipo, objEmail.iTipo)
    
    If Len(Trim(objEmail.sModelo)) > 0 Then
'        Mensagem.Enabled = False
        Modelo.Text = objEmail.sModelo
'    Else
'        Mensagem.Enabled = True
'        Modelo.Text = ""
    End If
    
    De.Text = objEmail.sDe
    NomeExibicao.Text = objEmail.sNomeExibicao
    
    Usuario.Text = objEmail.sUsuarioExclusivo
    Call Usuario_Validate(bSGECancelDummy)
    
    Traz_Email_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Email_Tela:

    Traz_Email_Tela = gErr

    Select Case gErr
    
        Case 187062
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187063)
    
    End Select
    
    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objEmail As ClassCobrancaEmailPadrao) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objEmail.lCodigo = StrParaLong(Codigo.Text)
    objEmail.sDescricao = DescReduzida.Text
    objEmail.iAtrasoAte = StrParaInt(AtrasoAte.Text)
    objEmail.iAtrasoDe = StrParaInt(AtrasoDe.Text)
    objEmail.sCC = Cc.Text
    objEmail.sAssunto = Assunto.Text
    objEmail.sMensagem = Mensagem.Text
    objEmail.sModelo = Modelo.Text
    objEmail.sAnexo = Anexo.Text
    objEmail.iTipo = Codigo_Extrai(Tipo.Text)
    objEmail.sDe = De.Text
    objEmail.sNomeExibicao = NomeExibicao.Text
    objEmail.sUsuarioExclusivo = Usuario.Text
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187064)
    
    End Select
    
    Exit Function

End Function

Function Critica_Dados(ByVal objEmail As ClassCobrancaEmailPadrao) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Dados

    If objEmail.iAtrasoDe > objEmail.iAtrasoAte Then gError 187065

    Critica_Dados = SUCESSO
    
    Exit Function

Erro_Critica_Dados:

    Critica_Dados = gErr

    Select Case gErr
    
        Case 187065
            Call Rotina_Erro(vbOKOnly, "ERRO_ATRASODE_MAIOR_ATRASOATE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187066)
    
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
    If lErro <> SUCESSO Then gError 187067
    
    For Each objFormulaFuncao In colFormulaFuncao
        Funcoes.AddItem objFormulaFuncao.sFuncaoCombo
    Next
    
    Carga_Combobox_Funcoes = SUCESSO

    Exit Function

Erro_Carga_Combobox_Funcoes:

    Carga_Combobox_Funcoes = gErr

    Select Case gErr

        Case 187067
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187068)

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
    If lErro <> SUCESSO Then gError 187069
    
    For Each objFormulaOperador In colFormulaOperador
        Operadores.AddItem objFormulaOperador.sOperadorCombo
    Next
    
    Carga_Combobox_Operadores = SUCESSO

    Exit Function

Erro_Carga_Combobox_Operadores:

    Carga_Combobox_Operadores = gErr

    Select Case gErr

        Case 187069
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187070)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Mnemonicos() As Long
'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.

Dim colMnemonico As New Collection
Dim objMnemonico As ClassMnemonicoCobrEmail
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Mnemonicos
        
    Mnemonicos.Enabled = True
    Mnemonicos.Clear
    
    'leitura dos mnemonicos no BD
    lErro = CF("MnemonicoCobrEmail_Le", colMnemonico)
    If lErro <> SUCESSO Then gError 187071

    For Each objMnemonico In colMnemonico
        Mnemonicos.AddItem objMnemonico.sMnemonicoCombo
    Next
    
    Carga_Combobox_Mnemonicos = SUCESSO

    Exit Function

Erro_Carga_Combobox_Mnemonicos:

    Carga_Combobox_Mnemonicos = gErr

    Select Case gErr

        Case 187071
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187072)

    End Select
    
    Exit Function

End Function

Private Function Regra_Valida(objControle As Object, Cancel As Boolean) As Long

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim colMnemonico As New Collection
Dim colSubRegras As New Collection
Dim sSubRegras As String
Dim iSubRegra As Integer

On Error GoTo Erro_Regra_Valida

    If Len(Trim(objControle.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
        
            lErro = CF("MnemonicoCobrEmail_Le", colMnemonico)
            If lErro <> SUCESSO Then gError 187074
            
            lErro = CF("Regra_Retorna_SubRegras", objControle.Text, colSubRegras)
            If lErro <> SUCESSO Then gError 187074

            For iSubRegra = 1 To colSubRegras.Count
            
                sSubRegras = colSubRegras.Item(iSubRegra)

                lErro = CF("Valida_Formula_WFW", sSubRegras, TIPO_TEXTO, iInicio, iTamanho, colMnemonico)
                If lErro <> SUCESSO Then gError 187075

            
            Next
                
        End If
        
    End If

    Regra_Valida = SUCESSO
    
    Exit Function
    
Erro_Regra_Valida:

    Regra_Valida = gErr
    
    Cancel = True
    
    Select Case gErr
    
        Case 187074, 187075
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187076)
        
    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 187078
    
    Call Limpa_Tela_Email

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 187078
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187079)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
'grava os dados da tela

Dim lErro As Long
Dim objEmail As New ClassCobrancaEmailPadrao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 187080
    If Len(Trim(DescReduzida.Text)) = 0 Then gError 187081
    If Len(Trim(Modelo.Text)) = 0 Then gError 200226
    
    If Codigo_Extrai(Tipo.Text) = TIPO_COBRANCAEMAILPADRAO_AVISO Or Codigo_Extrai(Tipo.Text) = TIPO_COBRANCAEMAILPADRAO_COBRANCA Then
        If Len(Trim(AtrasoDe.Text)) = 0 Then gError 187082
        If Len(Trim(AtrasoAte.Text)) = 0 Then gError 187083
    End If

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(objEmail)
    If lErro <> SUCESSO Then gError 187084
    
    lErro = Critica_Dados(objEmail)
    If lErro <> SUCESSO Then gError 187085

    lErro = Trata_Alteracao(objEmail, objEmail.lCodigo)
    If lErro <> SUCESSO Then gError 187086

    'Grava a etapa no Banco de Dados
    lErro = CF("CobrancaEmailPadrao_Grava", objEmail)
    If lErro <> SUCESSO Then gError 187087

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 187080
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAILCOBPADRAO_CODIGO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 187081
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAILCOBPADRAO_DESCRICAO_PREENCHIDO", gErr)
            DescReduzida.SetFocus
            
        Case 187082
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXAATRASO_NAO_PREENCHIDO", gErr)
            AtrasoDe.SetFocus

        Case 187083
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXAATRASO_NAO_PREENCHIDO", gErr)
            AtrasoAte.SetFocus
            
        Case 187084 To 187087

        Case 200226
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", gErr)
            Modelo.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187088)

    End Select

    Exit Function
    
End Function

Public Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objEmail As New ClassCobrancaEmailPadrao
    
On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 187089

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_COBRANCAEMAILPADRAO")
    If vbMsgRes = vbYes Then
    
        objEmail.lCodigo = Codigo.Text
    
        'exclui o modelo padrão de contabilização em questão
        lErro = CF("CobrancaEmailPadrao_Exclui", objEmail)
        If lErro <> SUCESSO Then gError 187090
    
        Call Limpa_Tela_Email
        
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 187089
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAILCOBPADRAO_CODIGO_PREENCHIDO", gErr)
            Codigo.SetFocus
            
        Case 187090
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187091)
        
    End Select

    Exit Sub
    
End Sub

Function Limpa_Tela_Email() As Long

    Call Limpa_Tela(Me)
    
    Call Combo_Seleciona_ItemData(Tipo, TIPO_COBRANCAEMAILPADRAO_AVISO)
    
    Usuario.ListIndex = -1
    Usuario.Text = ""
    
    Checkbox_Verifica_Sintaxe.Value = vbChecked
    
    Limpa_Tela_Email = SUCESSO
    
End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 187092

    Call Limpa_Tela_Email
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 187092
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187093)
        
    End Select
    
End Sub

Public Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Mnemonicos_Click()

Dim iPos As Integer
Dim lErro As Long
Dim lPos As Long
Dim objMnemonico As New ClassMnemonicoCobrEmail
Dim sMnemonico As String

On Error GoTo Erro_Mnemonicos_Click
    
    If Len(Mnemonicos.Text) > 0 Then
        
        objMnemonico.sMnemonicoCombo = Mnemonicos.Text
    
        'retorna os dados do mnemonico passado como parametro
        lErro = CF("MnemonicoCobrEmail_Le_Mnemonico", objMnemonico)
        If lErro <> SUCESSO And lErro <> 187116 Then gError 187094

        If lErro = 187116 Then gError 187095
        
        Descricao.Text = objMnemonico.sMnemonicoDesc
        
        lPos = InStr(1, Mnemonicos.Text, "(")
        If lPos = 0 Then
            sMnemonico = Mnemonicos.Text
        Else
            sMnemonico = Mid(Mnemonicos.Text, 1, lPos)
        End If
        
        lErro = Mnemonicos1(sMnemonico)
        If lErro <> SUCESSO Then gError 187096
        
    End If
    
    Exit Sub
    
Erro_Mnemonicos_Click:

    Select Case gErr
    
        Case 187094, 187096
    
        Case 187095
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", gErr, objMnemonico.sMnemonicoCombo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187097)
            
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
    If lErro <> SUCESSO And lErro <> 36098 Then gError 187099
    
    Descricao.Text = objFormulaOperador.sOperadorDesc
    
    Call Operadores1
    
    Exit Sub
    
Erro_Operadores_Click:

    Select Case gErr
    
        Case 187099
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187100)
            
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
    
'    If Not (Me.ActiveControl Is objControl) Then
'
'        If iPos >= Len(objControl.Text) Then
'            iTamanho = 0
'        Else
'            iTamanho = Len(objControl.Text) - iPos
'        End If
'        objControl.Text = Mid(objControl.Text, 1, iPos) & sTexto & Mid(objControl.Text, iPos + 1, iTamanho)
'
'    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Funcoes1(sFuncao As String) As Long

On Error GoTo Erro_Funcoes1

    Select Case iFocus
    
        Case FOCUS_ASSUNTO
            Call Posiciona_Texto_Tela(Assunto, sFuncao)
                    
        Case FOCUS_CC
            Call Posiciona_Texto_Tela(Cc, sFuncao)
                    
        Case FOCUS_MENSAGEM
            Call Posiciona_Texto_Tela(Mensagem, sFuncao)
                    
        Case FOCUS_ANEXO
            Call Posiciona_Texto_Tela(Anexo, sFuncao)
                    
    End Select
    
    Funcoes1 = SUCESSO
    
    Exit Function
    
Erro_Funcoes1:

    Funcoes1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187101)
            
    End Select
        
    Exit Function

End Function

Function Mnemonicos1(sMnemonico As String) As Long

Dim iPos As Integer

On Error GoTo Erro_Mnemonicos1

    Select Case iFocus
    
        Case FOCUS_ASSUNTO
            Call Posiciona_Texto_Tela(Assunto, sMnemonico)
                    
        Case FOCUS_CC
            Call Posiciona_Texto_Tela(Cc, sMnemonico)
                    
        Case FOCUS_MENSAGEM
            Call Posiciona_Texto_Tela(Mensagem, sMnemonico)
                    
        Case FOCUS_ANEXO
            Call Posiciona_Texto_Tela(Anexo, sMnemonico)
                    
    End Select

    Mnemonicos1 = SUCESSO
    
    Exit Function
    
Erro_Mnemonicos1:

    Mnemonicos1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187102)
            
    End Select
        
    Exit Function

End Function

Function Operadores1() As Long

Dim iPos As Integer

On Error GoTo Erro_Operadores1

    Select Case iFocus
    
        Case FOCUS_ASSUNTO
            Call Posiciona_Texto_Tela(Assunto, Operadores.Text)
                    
        Case FOCUS_CC
            Call Posiciona_Texto_Tela(Cc, Operadores.Text)
                    
        Case FOCUS_MENSAGEM
            Call Posiciona_Texto_Tela(Mensagem, Operadores.Text)
                    
        Case FOCUS_ANEXO
            Call Posiciona_Texto_Tela(Anexo, Operadores.Text)
                    
    End Select
     
    Operadores1 = SUCESSO
    
    Exit Function
    
Erro_Operadores1:

    Operadores1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187103)
            
    End Select
        
    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objEmail As New ClassCobrancaEmailPadrao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CobrancaEmailPadrao"

    'Lê os dados da Tela PedidoVenda
    objEmail.lCodigo = StrParaLong(Codigo.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objEmail.lCodigo, 0, "Codigo"
    'Filtros para o Sistema de Setas
    
    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187104)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objEmail As New ClassCobrancaEmailPadrao

On Error GoTo Erro_Tela_Preenche

    objEmail.lCodigo = colCampoValor.Item("Codigo").vValor

    If objEmail.lCodigo <> 0 Then
        lErro = Traz_Email_Tela(objEmail)
        If lErro <> SUCESSO Then gError 187105
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 187105

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187106)

    End Select

    Exit Function

End Function

Private Sub Assunto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Assunto_GotFocus()
    iFocus = FOCUS_ASSUNTO
End Sub

Private Sub Assunto_Validate(Cancel As Boolean)
    Call Regra_Valida(Assunto, Cancel)
End Sub

Private Sub Anexo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Anexo_GotFocus()
    iFocus = FOCUS_ANEXO
End Sub

Private Sub Anexo_Validate(Cancel As Boolean)
    Call Regra_Valida(Anexo, Cancel)
End Sub

Private Sub AtrasoAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AtrasoDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CC_GotFocus()
    iFocus = FOCUS_CC
End Sub

Private Sub CC_Validate(Cancel As Boolean)
    Call Regra_Valida(Cc, Cancel)
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DescReduzida_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Mensagem_GotFocus()
    iFocus = FOCUS_MENSAGEM
End Sub

Private Sub Mensagem_Validate(Cancel As Boolean)
    Call Regra_Valida(Mensagem, Cancel)
End Sub

Private Sub LabelCodigo_Click()

Dim objEmail As New ClassCobrancaEmailPadrao
Dim colSelecao As Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objEmail.lCodigo = StrParaLong(Codigo.Text)

    Call Chama_Tela("CobrancaEmailPadraoLista", colSelecao, objEmail, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objEmail As ClassCobrancaEmailPadrao

    Set objEmail = obj1

    Call Traz_Email_Tela(objEmail)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo lote disponível
    lErro = CF("CobrancaEmailPadrao_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 187107

    Codigo.PromptInclude = False
    Codigo.Text = lCodigo
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 187107
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187108)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurar_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    On Error GoTo Erro_BotaoProcurar_Click
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Html Files" & _
    "(*.html)|*.html"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    Modelo.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoProcurar_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Public Sub AtrasoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtrasoDe_Validate

    If Len(Trim(AtrasoDe.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica2(AtrasoDe.Text)
    If lErro <> SUCESSO Then gError 189402
   
    Exit Sub

Erro_AtrasoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 189402
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189403)

    End Select

    Exit Sub

End Sub

Public Sub AtrasoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtrasoAte_Validate

    If Len(Trim(AtrasoAte.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica2(AtrasoAte.Text)
    If lErro <> SUCESSO Then gError 189402
   
    Exit Sub

Erro_AtrasoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 189402
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189403)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
    If Codigo_Extrai(Tipo) = TIPO_COBRANCAEMAILPADRAO_COBRANCA Or Codigo_Extrai(Tipo) = TIPO_COBRANCAEMAILPADRAO_AVISO Then
        FrameAtraso.Visible = True
        FrameAtraso.Enabled = True
    Else
        FrameAtraso.Visible = False
        FrameAtraso.Enabled = False
        AtrasoDe.Text = ""
        AtrasoAte.Text = ""
    End If
    
End Sub

Private Sub Tipo_Click()
    iAlterado = REGISTRO_ALTERADO
    If Codigo_Extrai(Tipo) = TIPO_COBRANCAEMAILPADRAO_COBRANCA Or Codigo_Extrai(Tipo) = TIPO_COBRANCAEMAILPADRAO_AVISO Then
        FrameAtraso.Visible = True
        FrameAtraso.Enabled = True
    Else
        FrameAtraso.Visible = False
        FrameAtraso.Enabled = False
        AtrasoDe.Text = ""
        AtrasoAte.Text = ""
    End If
    
End Sub

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    Usuario.Clear

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 200098

    For Each objUsuarios In colUsuarios
        Usuario.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case 200098

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200099)

    End Select

    Exit Function

End Function

Private Sub Usuario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Usuario_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Usuario_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Usuario_Validate
    
    'Verifica se algum codigo está selecionado
    'If Usuario.ListIndex = -1 Then Exit Sub
    
    If Len(Trim(Usuario.Text)) > 0 Then
    
        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = Usuario.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError 200112
        
        If lErro <> SUCESSO Then gError 200113
        
    End If
    
    Exit Sub
    
Erro_Usuario_Validate:

    Cancel = True

    Select Case gErr
            
        Case 200112
        
        Case 200113 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200114)
    
    End Select
    
    Exit Sub
    
End Sub
