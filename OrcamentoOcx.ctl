VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OrcamentoOcx 
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   6390
   Begin VB.CommandButton BotaoConta 
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
      Height          =   510
      Left            =   4065
      TabIndex        =   21
      Top             =   2430
      Width           =   1605
   End
   Begin VB.CommandButton BotaoCcl 
      Caption         =   "Centros de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4065
      TabIndex        =   20
      Top             =   3435
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "OrcamentoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "OrcamentoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "OrcamentoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "OrcamentoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Exercicio 
      Height          =   315
      ItemData        =   "OrcamentoOcx.ctx":0994
      Left            =   1065
      List            =   "OrcamentoOcx.ctx":0996
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   1860
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   1470
      TabIndex        =   3
      Top             =   1950
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   3255
      Left            =   3375
      TabIndex        =   5
      Top             =   1665
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   5741
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
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
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   3255
      Left            =   3375
      TabIndex        =   6
      Top             =   1665
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   5741
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
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
   Begin MSMask.MaskEdBox Conta 
      Height          =   315
      Left            =   1065
      TabIndex        =   1
      Top             =   855
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridOrcamento 
      Height          =   3255
      Left            =   135
      TabIndex        =   4
      Top             =   1635
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   13
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSMask.MaskEdBox Ccl 
      Height          =   315
      Left            =   5115
      TabIndex        =   2
      Top             =   855
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   10
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
   Begin VB.Label LabelTvwCcls 
      AutoSize        =   -1  'True
      Caption         =   "Centros de Custo/Lucro:"
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
      Left            =   3375
      TabIndex        =   19
      Top             =   1410
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Exercicio:"
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
      Left            =   120
      TabIndex        =   18
      Top             =   255
      Width           =   855
   End
   Begin VB.Label LabelTvwContas 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   3375
      TabIndex        =   17
      Top             =   1410
      Width           =   2445
   End
   Begin VB.Label LabelConta 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
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
      Left            =   390
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   16
      Top             =   900
      Width           =   585
   End
   Begin VB.Label LabelCcl 
      AutoSize        =   -1  'True
      Caption         =   "Centro de Custo/Lucro:"
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
      Left            =   3045
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   900
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Orçamento por Período"
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
      TabIndex        =   14
      Top             =   1425
      Width           =   1995
   End
   Begin VB.Label TotalOrcamento 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1335
      TabIndex        =   13
      Top             =   4920
      Width           =   1635
   End
   Begin VB.Label LabelTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total:"
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
      Left            =   600
      TabIndex        =   12
      Top             =   4920
      Width           =   510
   End
End
Attribute VB_Name = "OrcamentoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTOrcamento
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub Ccl_GotFocus()
     Call objCT.Ccl_GotFocus
End Sub

Private Sub Conta_Change()
     Call objCT.Conta_Change
End Sub

Private Sub Ccl_Change()
     Call objCT.Ccl_Change
End Sub

Private Sub Conta_GotFocus()
     Call objCT.Conta_GotFocus
End Sub

Private Sub Conta_Validate(Cancel As Boolean)
     Call objCT.Conta_Validate(Cancel)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
     Call objCT.Ccl_Validate(Cancel)
End Sub

Private Sub Exercicio_Click()
     Call objCT.Exercicio_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub LabelConta_Click()
     Call objCT.LabelConta_Click
End Sub

Private Sub LabelCcl_Click()
     Call objCT.LabelCcl_Click
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)
     Call objCT.TvwContas_Expand(objNode)
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.TvwContas_NodeClick(Node)
End Sub

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.TvwCcls_NodeClick(Node)
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTOrcamento
    Set objCT.objUserControl = Me
End Sub

Private Sub Valor_Change()
     Call objCT.Valor_Change
End Sub

Private Sub Valor_GotFocus()
     Call objCT.Valor_GotFocus
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)
     Call objCT.Valor_KeyPress(KeyAscii)
End Sub

Private Sub Valor_Validate(Cancel As Boolean)
     Call objCT.Valor_Validate(Cancel)
End Sub

Private Sub GridOrcamento_Click()
     Call objCT.GridOrcamento_Click
End Sub

Private Sub GridOrcamento_GotFocus()
     Call objCT.GridOrcamento_GotFocus
End Sub

Private Sub GridOrcamento_EnterCell()
     Call objCT.GridOrcamento_EnterCell
End Sub

Private Sub GridOrcamento_LeaveCell()
     Call objCT.GridOrcamento_LeaveCell
End Sub

Private Sub GridOrcamento_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridOrcamento_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridOrcamento_KeyPress(KeyAscii As Integer)
     Call objCT.GridOrcamento_KeyPress(KeyAscii)
End Sub

Private Sub GridOrcamento_Validate(Cancel As Boolean)
     Call objCT.GridOrcamento_Validate(Cancel)
End Sub

Private Sub GridOrcamento_RowColChange()
     Call objCT.GridOrcamento_RowColChange
End Sub

Private Sub GridOrcamento_Scroll()
     Call objCT.GridOrcamento_Scroll
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros(Optional objOrcamento As ClassOrcamento) As Long
     Trata_Parametros = objCT.Trata_Parametros(objOrcamento)
End Function

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub



Private Sub LabelTvwCcls_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTvwCcls, Source, X, Y)
End Sub

Private Sub LabelTvwCcls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTvwCcls, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelTvwContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTvwContas, Source, X, Y)
End Sub

Private Sub LabelTvwContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTvwContas, Button, Shift, X, Y)
End Sub

Private Sub LabelConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelConta, Source, X, Y)
End Sub

Private Sub LabelConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelConta, Button, Shift, X, Y)
End Sub

Private Sub LabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcl, Source, X, Y)
End Sub

Private Sub LabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcl, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub TotalOrcamento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalOrcamento, Source, X, Y)
End Sub

Private Sub TotalOrcamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalOrcamento, Button, Shift, X, Y)
End Sub

Private Sub LabelTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotal, Source, X, Y)
End Sub

Private Sub LabelTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotal, Button, Shift, X, Y)
End Sub

Private Sub BotaoConta_Click()
     Call objCT.BotaoConta_Click
End Sub

Private Sub BotaoCcl_Click()
     Call objCT.BotaoCcl_Click
End Sub

