VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ConfiguracaoSetupOcx 
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   LockControls    =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   6510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   645
      Width           =   6000
      Begin VB.Frame Frame2 
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   615
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
         Begin VB.OptionButton LotePorPeriodo 
            Caption         =   "Por Período"
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
            Left            =   270
            TabIndex        =   1
            Top             =   270
            Width           =   1470
         End
         Begin VB.OptionButton LotePorExercicio 
            Caption         =   "Por Exercício"
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
            Left            =   240
            TabIndex        =   2
            Top             =   735
            Width           =   1530
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Left            =   3060
         TabIndex        =   15
         Top             =   1080
         Width           =   1890
         Begin VB.OptionButton DocPorPeriodo 
            Caption         =   "Por Período"
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
            Left            =   285
            TabIndex        =   3
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton DocPorExercicio 
            Caption         =   "Por Exercício"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   270
            TabIndex        =   4
            Top             =   720
            Width           =   1515
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Permite que você escolha como será feita a reinicialização da numeração dos seguintes campos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   195
         TabIndex        =   19
         Top             =   255
         Width           =   5565
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   2
      Left            =   150
      TabIndex        =   9
      Top             =   705
      Visible         =   0   'False
      Width           =   6015
      Begin VB.ComboBox Natureza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "ConfiguracaoSetupOcx.ctx":0000
         Left            =   2700
         List            =   "ConfiguracaoSetupOcx.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1725
         Width           =   2000
      End
      Begin VB.ComboBox Origem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1260
         Width           =   2000
      End
      Begin VB.ComboBox TipoConta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "ConfiguracaoSetupOcx.ctx":0004
         Left            =   2700
         List            =   "ConfiguracaoSetupOcx.ctx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   780
         Width           =   2000
      End
      Begin VB.Label Nat 
         AutoSize        =   -1  'True
         Caption         =   "Natureza:"
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
         Left            =   1380
         TabIndex        =   20
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
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
         Left            =   1395
         TabIndex        =   21
         Top             =   1305
         Width           =   660
      End
      Begin VB.Label TipoDaConta 
         AutoSize        =   -1  'True
         Caption         =   "Tipo da Conta:"
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
         Left            =   1320
         TabIndex        =   22
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Valores Iniciais dos Campos nas Telas em que aparecem:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   23
         Top             =   420
         Width           =   5310
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(1)"
      Height          =   2700
      Index           =   1
      Left            =   285
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   5715
      Begin VB.Frame Frame4 
         Caption         =   "Utilização de Centro de Custo/Centro de Lucro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   5475
         Begin VB.OptionButton CclExtra 
            Caption         =   "Utiliza Centro de Custo/Centro de Lucro Extra Contábil"
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
            Left            =   315
            TabIndex        =   8
            Top             =   1320
            Width           =   5115
         End
         Begin VB.OptionButton CclContabil 
            Caption         =   "Utiliza Centro de Custo/Centro de Lucro Contábil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   345
            TabIndex        =   7
            Top             =   825
            Width           =   4515
         End
         Begin VB.OptionButton SemCcl 
            Caption         =   "Não utiliza Centro de Custo/Centro de Lucro"
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
            Left            =   330
            TabIndex        =   6
            Top             =   465
            Width           =   4245
         End
      End
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3300
      Picture         =   "ConfiguracaoSetupOcx.ctx":0008
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3705
      Width           =   975
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1875
      Picture         =   "ConfiguracaoSetupOcx.ctx":010A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3705
      Width           =   975
   End
   Begin MSComctlLib.TabStrip Opcoes 
      Height          =   3405
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   6006
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicialização"
            Object.ToolTipText     =   "Indica como serão reinicializadas as numerações de alguns campos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Centro de Custo/Lucro"
            Object.ToolTipText     =   "Utilização de centro de custo/centro de lucro"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Valores Iniciais"
            Object.ToolTipText     =   "Valores com que os campos serão inicializados"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ConfiguracaoSetupOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Responsavel: Mario
'Revisado em 20/8/98

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer
Dim iFrameAtual As Integer

'Constantes públicas dos tabs
Private Const TAB_Inicializacao = 1
Private Const TAB_Ccl = 2
Private Const TAB_ValoresIniciais = 3

Function Trata_Parametros() As Long
    
    iAlterado = 0

    Trata_Parametros = SUCESSO
    
End Function

Private Sub BotaoCancela_Click()

    Unload Me
    
End Sub

Private Sub CclContabil_Click()
  
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclExtra_Click()
  
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DocPorExercicio_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DocPorPeriodo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objConfiguracao As New ClassConfiguracao
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
        
    'Le o registro da tabela Configuracao
    lErro = CF("Configuracao_Le",objConfiguracao)
    If lErro <> SUCESSO Then Error 12172

    'marca as opcoes na parte do Lote item Inicializacao
    If objConfiguracao.iLotePorPeriodo = LOTE_INICIALIZADO_POR_PERIODO Then
        LotePorPeriodo.Value = True
        LotePorExercicio.Value = False
    Else
        LotePorPeriodo.Value = False
        LotePorExercicio.Value = True
    End If
    
    'marca as opcoes na parte do Documento no item Inicializacao
    If objConfiguracao.iDocPorPeriodo = DOC_INICIALIZADO_POR_PERIODO Then
        DocPorPeriodo.Value = True
        DocPorExercicio.Value = False
    Else
        DocPorPeriodo.Value = False
        DocPorExercicio.Value = True
    End If

    'marca a opcao referente no item Centro de Custo/Lucro
    Select Case objConfiguracao.iUsoCcl
    
        Case CCL_NAO_USA
                SemCcl.Value = True
        Case CCL_USA_CONTABIL
                CclContabil = True
        Case CCL_USA_EXTRACONTABIL
                CclExtra = True
    End Select
    
    'inicializar os tipos de conta
    For iIndice = 1 To gobjColTipoConta.Count
        TipoConta.AddItem gobjColTipoConta.Item(iIndice).sDescricao
    Next
    
    
    'inicializar as naturezas de conta
    For iIndice = 1 To gobjColNaturezaConta.Count
        Natureza.AddItem gobjColNaturezaConta.Item(iIndice).sDescricao
    Next
    
    'mostra o TipoConta que esta na tabela Configuracao
    For iIndice = 0 To gobjColTipoConta.Count - 1
        TipoConta.ListIndex = iIndice
        If TipoConta.Text = gobjColTipoConta.Descricao(objConfiguracao.iTipoContaDefault) Then Exit For
    Next

    'mostra a Natureza que esta na tabela Confuguracao
    For iIndice = 0 To gobjColNaturezaConta.Count - 1
        Natureza.ListIndex = iIndice
        If Natureza.Text = gobjColNaturezaConta.Descricao(objConfiguracao.iNaturezaDefault) Then Exit For
    Next
    
    iAlterado = 0
    iFrameAtual = 0

    lErro_Chama_Tela = SUCESSO
            
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 12172
                        
        Case Else
            
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154674)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Private Sub Frame1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

End Sub

Private Sub LotePorExercicio_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LotePorPeriodo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Opcoes_Click()

    If Opcoes.SelectedItem.Index - 1 <> iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcoes, Me) <> SUCESSO Then Exit Sub
        
        Frame1(Opcoes.SelectedItem.Index - 1).Visible = True
        Frame1(iFrameAtual).Visible = False
        iFrameAtual = Opcoes.SelectedItem.Index - 1
                
        Select Case iFrameAtual
        
            Case TAB_Inicializacao
                Parent.HelpContextID = IDH_CONFIGURACAO_SETUP_ID
                
            Case TAB_Ccl
                Parent.HelpContextID = IDH_CONFIGURACAO_SETUP_CCL
                        
            Case TAB_ValoresIniciais
                Parent.HelpContextID = IDH_CONFIGURACAO_SETUP_VALORES_INICIAIS
                
        End Select
        
    End If

End Sub

Private Sub SemCcl_Click()
  
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoConta_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoConta_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub Natureza_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Natureza_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub BotaoOK_Click()
'faz as gravacoes das configuracoes

    Call Gravar_Registro
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Function Leitura_Configuracao(objConfiguracao As ClassConfiguracao) As Long
'faz a leitura das marcacoes da tela de ConfiguracaoSetup

    'le a marcacao do Lote
    If LotePorPeriodo.Value Then
        objConfiguracao.iLotePorPeriodo = 1
    Else
        objConfiguracao.iLotePorPeriodo = 0
    End If
    
    'le a marcacao do Documento
    If DocPorPeriodo.Value Then
        objConfiguracao.iDocPorPeriodo = 1
    Else
        objConfiguracao.iDocPorPeriodo = 0
    End If
    
    'le a marcacao do Centro de Custo/Lucro
    If SemCcl.Value Then
        objConfiguracao.iUsoCcl = 0
    ElseIf CclContabil.Value Then
        objConfiguracao.iUsoCcl = 1
    ElseIf CclExtra.Value Then
        objConfiguracao.iUsoCcl = 2
    End If
    
    objConfiguracao.iTipoContaDefault = gobjColTipoConta.TipoConta(TipoConta.Text)
    objConfiguracao.iNaturezaDefault = gobjColNaturezaConta.NaturezaConta(Natureza.Text)

    Leitura_Configuracao = SUCESSO

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objConfiguracao As New ClassConfiguracao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    Call Leitura_Configuracao(objConfiguracao)
        
    'Grava os registros na tabela Configuracao com os dados de objConfiguracao
    lErro = CF("ConfiguracaoSetup_Grava",objConfiguracao)
    If lErro <> SUCESSO Then Error 12173
    
    iAlterado = 0
    
    Unload Me
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function
    
Erro_Gravar_Registro:
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 12173
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154675)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONFIGURACAO_SETUP_ID
    Set Form_Load_Ocx = Me
    Caption = "Configuração Setup"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConfiguracaoSetup"
    
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




Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Nat_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Nat, Source, X, Y)
End Sub

Private Sub Nat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Nat, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub TipoDaConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoDaConta, Source, X, Y)
End Sub

Private Sub TipoDaConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoDaConta, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub


Private Sub Opcoes_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcoes)
End Sub

