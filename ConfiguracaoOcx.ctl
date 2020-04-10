VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ConfiguracaoOcx 
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LockControls    =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   6495
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   4
      Left            =   255
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   6000
      Begin VB.CheckBox HistoricoObrigatorio 
         Caption         =   "Obriga o preenchimento do histórico na contabilidade."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   525
         TabIndex        =   34
         Top             =   210
         Width           =   5130
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   2565
      Index           =   3
      Left            =   180
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   6015
      Begin MSMask.MaskEdBox ContaResultado 
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   510
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSComctlLib.TreeView TvwContas 
         Height          =   2265
         Left            =   3840
         TabIndex        =   16
         Top             =   285
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3995
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox ContaTransferencia 
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   1155
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaProducao 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   1830
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conta de Resultado:"
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
         Left            =   345
         TabIndex        =   22
         Top             =   570
         Width           =   1770
      End
      Begin VB.Label LabelContas 
         Caption         =   "Plano de Contas"
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
         Left            =   3855
         TabIndex        =   21
         Top             =   60
         Width           =   2325
      End
      Begin VB.Label Label2 
         Caption         =   "Conta de Transferência de Valores Entre Filiais:"
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
         Height          =   420
         Left            =   75
         TabIndex        =   20
         Top             =   1095
         Width           =   2040
      End
      Begin VB.Label LabelContaProducao 
         AutoSize        =   -1  'True
         Caption         =   "Conta de Produção:"
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
         Left            =   390
         TabIndex        =   19
         ToolTipText     =   "Conta contábil de aplicação"
         Top             =   1860
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2565
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   6015
      Begin VB.ComboBox TipoConta 
         Height          =   315
         ItemData        =   "ConfiguracaoOcx.ctx":0000
         Left            =   2700
         List            =   "ConfiguracaoOcx.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1020
         Width           =   2000
      End
      Begin VB.ComboBox Natureza 
         Height          =   315
         ItemData        =   "ConfiguracaoOcx.ctx":0004
         Left            =   2700
         List            =   "ConfiguracaoOcx.ctx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1755
         Width           =   2000
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
         TabIndex        =   25
         Top             =   420
         Width           =   5310
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
         TabIndex        =   24
         Top             =   1080
         Width           =   1275
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
         Left            =   1755
         TabIndex        =   23
         Top             =   1800
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(1)"
      Height          =   2700
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   960
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
         Left            =   225
         TabIndex        =   30
         Top             =   795
         Width           =   5475
         Begin VB.OptionButton SemCcl 
            Caption         =   "Não utiliza Centro de Custo/Centro de Lucro"
            Enabled         =   0   'False
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
         Begin VB.OptionButton CclContabil 
            Caption         =   "Utiliza Centro de Custo/Centro de Lucro Contábil"
            Enabled         =   0   'False
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
            Left            =   330
            TabIndex        =   7
            Top             =   825
            Width           =   4515
         End
         Begin VB.OptionButton CclExtra 
            Caption         =   "Utiliza Centro de Custo/Centro de Lucro Extra Contábil"
            Enabled         =   0   'False
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
            Left            =   330
            TabIndex        =   8
            Top             =   1320
            Width           =   5115
         End
      End
      Begin VB.Label Label12 
         Caption         =   "Atenção! A alteração desta opção só pode ser realizada no momento em que a empresa está sendo criada no sistema."
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   300
         TabIndex        =   31
         Top             =   270
         Width           =   5355
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2595
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   960
      Width           =   6000
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
         Height          =   1065
         Index           =   0
         Left            =   3060
         TabIndex        =   27
         Top             =   1425
         Width           =   1890
         Begin VB.OptionButton DocPorExercicio 
            Caption         =   "Por Exercício"
            Enabled         =   0   'False
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
         Begin VB.OptionButton DocPorPeriodo 
            Caption         =   "Por Período"
            Enabled         =   0   'False
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
            TabIndex        =   3
            Top             =   285
            Width           =   1455
         End
      End
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
         Height          =   1035
         Index           =   0
         Left            =   615
         TabIndex        =   26
         Top             =   1440
         Width           =   1935
         Begin VB.OptionButton LotePorExercicio 
            Caption         =   "Por Exercício"
            Enabled         =   0   'False
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
            Top             =   705
            Width           =   1530
         End
         Begin VB.OptionButton LotePorPeriodo 
            Caption         =   "Por Período"
            Enabled         =   0   'False
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
            TabIndex        =   1
            Top             =   270
            Width           =   1470
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Atenção! A alteração desta opção só pode ser feita no momento da criação da empresa no sistema."
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   210
         TabIndex        =   29
         Top             =   795
         Width           =   5595
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
         Left            =   180
         TabIndex        =   28
         Top             =   255
         Width           =   5565
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
      Left            =   3315
      Picture         =   "ConfiguracaoOcx.ctx":0008
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3975
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
      Left            =   1905
      Picture         =   "ConfiguracaoOcx.ctx":010A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3975
      Width           =   975
   End
   Begin MSComctlLib.TabStrip Opcoes 
      Height          =   3675
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   6482
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
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
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Outros"
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
Attribute VB_Name = "ConfiguracaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTConfiguracao
Attribute objCT.VB_VarHelpID = -1

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub BotaoOK_Click()
     Call objCT.BotaoOK_Click
End Sub

Private Sub BotaoCancela_Click()
     Call objCT.BotaoCancela_Click
End Sub

Private Sub ContaResultado_Change()
     Call objCT.ContaResultado_Change
End Sub

Private Sub ContaResultado_Validate(Cancel As Boolean)
     Call objCT.ContaResultado_Validate(Cancel)
End Sub

Private Sub ContaTransferencia_Change()
     Call objCT.ContaTransferencia_Change
End Sub

Private Sub ContaTransferencia_Validate(Cancel As Boolean)
     Call objCT.ContaTransferencia_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub Opcoes_Click()
     Call objCT.Opcoes_Click
End Sub

Private Sub TipoConta_Change()
     Call objCT.TipoConta_Change
End Sub

Private Sub TipoConta_Click()
     Call objCT.TipoConta_Click
End Sub

Private Sub Natureza_Change()
     Call objCT.Natureza_Change
End Sub

Private Sub Natureza_Click()
     Call objCT.Natureza_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)
     Call objCT.TvwContas_Expand(objNode)
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.TvwContas_NodeClick(Node)
End Sub

Private Sub ContaProducao_Change()
     Call objCT.ContaProducao_Change
End Sub

Private Sub ContaProducao_Validate(Cancel As Boolean)
     Call objCT.ContaProducao_Validate(Cancel)
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
            Set objCT.objUserControl = Nothing
            Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Private Sub UserControl_Initialize()
    Set objCT = New CTConfiguracao
    Set objCT.objUserControl = Me
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property


Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelContaProducao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaProducao, Source, X, Y)
End Sub

Private Sub LabelContaProducao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaProducao, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub TipoDaConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoDaConta, Source, X, Y)
End Sub

Private Sub TipoDaConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoDaConta, Button, Shift, X, Y)
End Sub

Private Sub Nat_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Nat, Source, X, Y)
End Sub

Private Sub Nat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Nat, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub


Private Sub Opcoes_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcoes)
End Sub

