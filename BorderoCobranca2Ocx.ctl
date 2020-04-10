VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl BorderoCobranca2Ocx 
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   ScaleHeight     =   4995
   ScaleWidth      =   9165
   Begin VB.ComboBox Ordenacao 
      Height          =   315
      ItemData        =   "BorderoCobranca2Ocx.ctx":0000
      Left            =   1200
      List            =   "BorderoCobranca2Ocx.ctx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   150
      Width           =   2820
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   3270
      ScaleHeight     =   495
      ScaleWidth      =   2685
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2745
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "BorderoCobranca2Ocx.ctx":0042
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   150
         Picture         =   "BorderoCobranca2Ocx.ctx":01C0
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   90
         Width           =   885
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1117
         Picture         =   "BorderoCobranca2Ocx.ctx":091E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.TextBox Tipo 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   180
      TabIndex        =   12
      Top             =   705
      Width           =   690
   End
   Begin VB.CommandButton BotaoDocOriginal 
      Height          =   690
      Left            =   1185
      Picture         =   "BorderoCobranca2Ocx.ctx":10B0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4185
      Width           =   1740
   End
   Begin VB.CommandButton BotaoDesmarcar 
      Caption         =   "Desmarcar Todas"
      Height          =   585
      Left            =   3750
      Picture         =   "BorderoCobranca2Ocx.ctx":3FC6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1440
   End
   Begin VB.CommandButton BotaoMarcar 
      Caption         =   "Marcar Todas"
      Height          =   585
      Left            =   3750
      Picture         =   "BorderoCobranca2Ocx.ctx":51A8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2940
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total"
      Height          =   960
      Left            =   1185
      TabIndex        =   14
      Top             =   2985
      Width           =   2250
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
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
         Left            =   165
         TabIndex        =   15
         Top             =   270
         Width           =   540
      End
      Begin VB.Label QtdParcelas 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   780
         TabIndex        =   16
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label6 
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
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   615
         Width           =   510
      End
      Begin VB.Label TotalParcelas 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   780
         TabIndex        =   18
         Top             =   585
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecionados"
      Height          =   960
      Left            =   5460
      TabIndex        =   11
      Top             =   3000
      Width           =   2250
      Begin VB.Label QtdParcelasSelecionadas 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   765
         TabIndex        =   19
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
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
         TabIndex        =   20
         Top             =   300
         Width           =   540
      End
      Begin VB.Label TotalParcelasSelecionadas 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   765
         TabIndex        =   21
         Top             =   555
         Width           =   1275
      End
      Begin VB.Label Label5 
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
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   585
         Width           =   510
      End
   End
   Begin VB.CheckBox CheckIncluir 
      Height          =   255
      Left            =   315
      TabIndex        =   0
      Top             =   375
      Width           =   495
   End
   Begin MSMask.MaskEdBox DataVencto 
      Height          =   225
      Left            =   7455
      TabIndex        =   7
      Top             =   345
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   6540
      TabIndex        =   6
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   225
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid GridBorderoCobranca2 
      Height          =   2250
      Left            =   105
      TabIndex        =   8
      Top             =   600
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   3969
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSMask.MaskEdBox Parcela 
      Height          =   225
      Left            =   5700
      TabIndex        =   5
      Top             =   360
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox NumTitulo 
      Height          =   225
      Left            =   4620
      TabIndex        =   4
      Top             =   150
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox Filial 
      Height          =   225
      Left            =   3900
      TabIndex        =   3
      Top             =   360
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox Nome 
      Height          =   225
      Left            =   1515
      TabIndex        =   2
      Top             =   360
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ordenação:"
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
      Left            =   165
      TabIndex        =   28
      Top             =   210
      Width           =   1005
   End
End
Attribute VB_Name = "BorderoCobranca2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTBorderoCobranca2
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTBorderoCobranca2
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoDesmarcar_Click()
     Call objCT.BotaoDesmarcar_Click
End Sub

Private Sub BotaoDocOriginal_Click()
     Call objCT.BotaoDocOriginal_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoMarcar_Click()
     Call objCT.BotaoMarcar_Click
End Sub

Private Sub BotaoSeguir_Click()
     Call objCT.BotaoSeguir_Click
End Sub

Private Sub BotaoVoltar_Click()
     Call objCT.BotaoVoltar_Click
End Sub

Private Sub CheckIncluir_Click()
     Call objCT.CheckIncluir_Click
End Sub

Private Sub CheckIncluir_GotFocus()
     Call objCT.CheckIncluir_GotFocus
End Sub

Private Sub CheckIncluir_KeyPress(KeyAscii As Integer)
     Call objCT.CheckIncluir_KeyPress(KeyAscii)
End Sub

Private Sub CheckIncluir_Validate(Cancel As Boolean)
     Call objCT.CheckIncluir_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub GridBorderoCobranca2_Click()
     Call objCT.GridBorderoCobranca2_Click
End Sub

Private Sub GridBorderoCobranca2_GotFocus()
     Call objCT.GridBorderoCobranca2_GotFocus
End Sub

Private Sub GridBorderoCobranca2_EnterCell()
     Call objCT.GridBorderoCobranca2_EnterCell
End Sub

Private Sub GridBorderoCobranca2_LeaveCell()
     Call objCT.GridBorderoCobranca2_LeaveCell
End Sub

Private Sub GridBorderoCobranca2_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridBorderoCobranca2_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridBorderoCobranca2_KeyPress(KeyAscii As Integer)
     Call objCT.GridBorderoCobranca2_KeyPress(KeyAscii)
End Sub

Private Sub GridBorderoCobranca2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid
    Call objCT.GridBorderoCobranca2_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub GridBorderoCobranca2_Validate(Cancel As Boolean)
     Call objCT.GridBorderoCobranca2_Validate(Cancel)
End Sub

Private Sub GridBorderoCobranca2_RowColChange()
     Call objCT.GridBorderoCobranca2_RowColChange
End Sub

Private Sub GridBorderoCobranca2_Scroll()
     Call objCT.GridBorderoCobranca2_Scroll
End Sub

Function Trata_Parametros(Optional objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long
     Trata_Parametros = objCT.Trata_Parametros(objBorderoCobrancaEmissao)
End Function

Private Sub QtdParcelasSelecionadas_Change()
     Call objCT.QtdParcelasSelecionadas_Change
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Private Sub QtdParcelas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdParcelas, Source, X, Y)
End Sub
Private Sub QtdParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdParcelas, Button, Shift, X, Y)
End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
Private Sub TotalParcelas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalParcelas, Source, X, Y)
End Sub
Private Sub TotalParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalParcelas, Button, Shift, X, Y)
End Sub
Private Sub QtdParcelasSelecionadas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdParcelasSelecionadas, Source, X, Y)
End Sub
Private Sub QtdParcelasSelecionadas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdParcelasSelecionadas, Button, Shift, X, Y)
End Sub
Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub
Private Sub TotalParcelasSelecionadas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalParcelasSelecionadas, Source, X, Y)
End Sub
Private Sub TotalParcelasSelecionadas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalParcelasSelecionadas, Button, Shift, X, Y)
End Sub
Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub
Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_UnLoad(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
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

Private Sub Ordenacao_Click()
    Call objCT.Ordenacao_Click
End Sub

