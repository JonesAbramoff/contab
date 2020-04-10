VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl DocAutoOcx 
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   KeyPreview      =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   9420
   Begin VB.CheckBox Gerencial 
      Height          =   210
      Left            =   4680
      TabIndex        =   32
      Tag             =   "1"
      Top             =   2280
      Width           =   870
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2040
      Picture         =   "DocAutoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   135
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7200
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "DocAutoOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "DocAutoOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "DocAutoOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "DocAutoOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListaHistorico 
      Height          =   3570
      Left            =   6930
      TabIndex        =   10
      Top             =   1290
      Width           =   2430
   End
   Begin VB.TextBox Historico 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5295
      MaxLength       =   150
      TabIndex        =   8
      Top             =   1755
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descrição do Elemento Selecionado"
      Height          =   1230
      Left            =   225
      TabIndex        =   18
      Top             =   3600
      Width           =   6315
      Begin VB.Label CclDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1740
         TabIndex        =   22
         Top             =   750
         Visible         =   0   'False
         Width           =   3720
      End
      Begin VB.Label ContaDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1740
         TabIndex        =   21
         Top             =   345
         Width           =   3720
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1125
         TabIndex        =   20
         Top             =   360
         Width           =   570
      End
      Begin VB.Label CclLabel 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo:"
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
         Left            =   255
         TabIndex        =   19
         Top             =   765
         Visible         =   0   'False
         Width           =   1440
      End
   End
   Begin MSMask.MaskEdBox SeqContraPartida 
      Height          =   225
      Left            =   4860
      TabIndex        =   7
      Top             =   1770
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Debito 
      Height          =   225
      Left            =   2415
      TabIndex        =   5
      Top             =   1770
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Credito 
      Height          =   225
      Left            =   3660
      TabIndex        =   6
      Top             =   1755
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Ccl 
      Height          =   225
      Left            =   1740
      TabIndex        =   4
      Top             =   1695
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
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
   Begin MSMask.MaskEdBox Conta 
      Height          =   225
      Left            =   570
      TabIndex        =   3
      Top             =   1695
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
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
   Begin MSMask.MaskEdBox Documento 
      Height          =   315
      Left            =   1335
      TabIndex        =   0
      Top             =   120
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1335
      TabIndex        =   2
      Top             =   510
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   40
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
   Begin MSFlexGridLib.MSFlexGrid GridDocAuto 
      Height          =   1860
      Left            =   75
      TabIndex        =   9
      Top             =   1275
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSComctlLib.TreeView TvwCcls 
      Height          =   3570
      Left            =   6945
      TabIndex        =   11
      Top             =   1290
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   6297
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
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   3570
      Left            =   6930
      TabIndex        =   12
      Top             =   1290
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   6297
      _Version        =   393217
      Indentation     =   511
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
   Begin VB.Label LabelHistorico 
      AutoSize        =   -1  'True
      Caption         =   "Históricos"
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
      Left            =   6930
      TabIndex        =   31
      Top             =   1065
      Width           =   2385
   End
   Begin VB.Label LabelCCL 
      AutoSize        =   -1  'True
      Caption         =   "Centros de Custo/Lucro"
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
      Left            =   6930
      TabIndex        =   30
      Top             =   1065
      Width           =   2040
   End
   Begin VB.Label LabelPlanoConta 
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
      Height          =   240
      Left            =   6930
      TabIndex        =   29
      Top             =   1065
      Width           =   2385
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   -45
      X2              =   9585
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
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
      Height          =   240
      Left            =   630
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   28
      Top             =   180
      Width           =   660
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   570
      Width           =   945
   End
   Begin VB.Label TotalCredito 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3135
      TabIndex        =   26
      Top             =   3165
      Width           =   1155
   End
   Begin VB.Label TotalDebito 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1875
      TabIndex        =   25
      Top             =   3165
      Width           =   1155
   End
   Begin VB.Label LabelTotais 
      AutoSize        =   -1  'True
      Caption         =   "Totais:"
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
      Left            =   1170
      TabIndex        =   24
      Top             =   3180
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lançamentos"
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
      Left            =   120
      TabIndex        =   23
      Top             =   1065
      Width           =   1140
   End
End
Attribute VB_Name = "DocAutoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTDocAuto
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros(Optional objDocAuto As ClassDocAuto) As Long
     Trata_Parametros = objCT.Trata_Parametros(objDocAuto)
End Function

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub Descricao_Change()
     Call objCT.Descricao_Change
End Sub

Private Sub Conta_Change()
     Call objCT.Conta_Change
End Sub

Private Sub Ccl_Change()
     Call objCT.Ccl_Change
End Sub

Private Sub Credito_Change()
     Call objCT.Credito_Change
End Sub

Private Sub Debito_Change()
     Call objCT.Debito_Change
End Sub

Private Sub Documento_GotFocus()
     Call objCT.Documento_GotFocus
End Sub

Private Sub Historico_Change()
     Call objCT.Historico_Change
End Sub

Private Sub GridDocAuto_LeaveCell()
     Call objCT.GridDocAuto_LeaveCell
End Sub

Private Sub GridDocAuto_EnterCell()
     Call objCT.GridDocAuto_EnterCell
End Sub

Private Sub GridDocAuto_Click()
     Call objCT.GridDocAuto_Click
End Sub

Private Sub GridDocAuto_GotFocus()
     Call objCT.GridDocAuto_GotFocus
End Sub

Private Sub GridDocAuto_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridDocAuto_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridDocAuto_KeyPress(KeyAscii As Integer)
     Call objCT.GridDocAuto_KeyPress(KeyAscii)
End Sub

Private Sub GridDocAuto_Validate(Cancel As Boolean)
     Call objCT.GridDocAuto_Validate(Cancel)
End Sub

Private Sub Label3_Click()
     Call objCT.Label3_Click
End Sub

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.TvwCcls_NodeClick(Node)
End Sub

Private Sub ListaHistorico_DblClick()
     Call objCT.ListaHistorico_DblClick
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)
     Call objCT.TvwContas_Expand(objNode)
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.TvwContas_NodeClick(Node)
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub Conta_GotFocus()
     Call objCT.Conta_GotFocus
End Sub

Private Sub Conta_KeyPress(KeyAscii As Integer)
     Call objCT.Conta_KeyPress(KeyAscii)
End Sub

Private Sub Conta_Validate(Cancel As Boolean)
     Call objCT.Conta_Validate(Cancel)
End Sub

Private Sub Ccl_GotFocus()
     Call objCT.Ccl_GotFocus
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)
     Call objCT.Ccl_KeyPress(KeyAscii)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
     Call objCT.Ccl_Validate(Cancel)
End Sub

Private Sub Credito_GotFocus()
     Call objCT.Credito_GotFocus
End Sub

Private Sub Credito_KeyPress(KeyAscii As Integer)
     Call objCT.Credito_KeyPress(KeyAscii)
End Sub

Private Sub Credito_Validate(Cancel As Boolean)
     Call objCT.Credito_Validate(Cancel)
End Sub

Private Sub Debito_GotFocus()
     Call objCT.Debito_GotFocus
End Sub

Private Sub Debito_KeyPress(KeyAscii As Integer)
     Call objCT.Debito_KeyPress(KeyAscii)
End Sub

Private Sub Debito_Validate(Cancel As Boolean)
     Call objCT.Debito_Validate(Cancel)
End Sub

Private Sub SeqContraPartida_Change()
     Call objCT.SeqContraPartida_Change
End Sub

Private Sub SeqContraPartida_GotFocus()
     Call objCT.SeqContraPartida_GotFocus
End Sub

Private Sub SeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.SeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub SeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.SeqContraPartida_Validate(Cancel)
End Sub

Private Sub Documento_Change()
     Call objCT.Documento_Change
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Historico_GotFocus()
     Call objCT.Historico_GotFocus
End Sub

Private Sub Historico_KeyPress(KeyAscii As Integer)
     Call objCT.Historico_KeyPress(KeyAscii)
End Sub

Private Sub Historico_Validate(Cancel As Boolean)
     Call objCT.Historico_Validate(Cancel)
End Sub

Private Sub GridDocAuto_RowColChange()
     Call objCT.GridDocAuto_RowColChange
End Sub

Private Sub GridDocAuto_Scroll()
     Call objCT.GridDocAuto_Scroll
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
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
    Set objCT = New CTDocAuto
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub



Private Sub CclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclDescricao, Source, X, Y)
End Sub

Private Sub CclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclDescricao, Button, Shift, X, Y)
End Sub

Private Sub ContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaDescricao, Source, X, Y)
End Sub

Private Sub ContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelHistorico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHistorico, Source, X, Y)
End Sub

Private Sub LabelHistorico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHistorico, Button, Shift, X, Y)
End Sub

Private Sub LabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcl, Source, X, Y)
End Sub

Private Sub LabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcl, Button, Shift, X, Y)
End Sub

Private Sub LabelPlanoConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPlanoConta, Source, X, Y)
End Sub

Private Sub LabelPlanoConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPlanoConta, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub TotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCredito, Source, X, Y)
End Sub

Private Sub TotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCredito, Button, Shift, X, Y)
End Sub

Private Sub TotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalDebito, Source, X, Y)
End Sub

Private Sub TotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalDebito, Button, Shift, X, Y)
End Sub

Private Sub LabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTotais, Source, X, Y)
End Sub

Private Sub LabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTotais, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub


