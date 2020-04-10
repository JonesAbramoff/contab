VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ChequesPag3Ocx 
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   ScaleHeight     =   5655
   ScaleWidth      =   8040
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   2685
      ScaleHeight     =   495
      ScaleWidth      =   2610
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5010
      Width           =   2670
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   75
         Picture         =   "ChequesPag3Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2115
         Picture         =   "ChequesPag3Ocx.ctx":075E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1110
         Picture         =   "ChequesPag3Ocx.ctx":08DC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controle de Impressão de Cheques"
      Height          =   1740
      Left            =   495
      TabIndex        =   15
      Top             =   3180
      Width           =   7065
      Begin VB.CommandButton BotaoImprimirAPartir 
         Caption         =   "Imprimir a Partir do Cheque Selecionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   450
         TabIndex        =   10
         Top             =   1215
         Width           =   6195
      End
      Begin VB.CommandButton BotaoConfigurarImpressao 
         Caption         =   "Configurar Impressão..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   420
         TabIndex        =   6
         Top             =   315
         Width           =   3015
      End
      Begin VB.CommandButton BotaoImprimirSelecao 
         Caption         =   "Imprimir os  Cheques Selecionados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   435
         TabIndex        =   8
         Top             =   795
         Width           =   3360
      End
      Begin VB.CommandButton BotaoImprimirTeste 
         Caption         =   "Imprimir Teste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3600
         TabIndex        =   7
         Top             =   315
         Width           =   3015
      End
      Begin VB.CommandButton BotaoImprimirTudo 
         Caption         =   "Imprimir Todos os Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3930
         TabIndex        =   9
         Top             =   780
         Width           =   2700
      End
   End
   Begin VB.CommandButton BotaoNumAuto 
      Caption         =   "Gerar numeração automática dos Cheques abaixo do Cheque selecionado"
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
      Left            =   450
      TabIndex        =   5
      Top             =   2685
      Width           =   7140
   End
   Begin VB.CheckBox Atualizar 
      BackColor       =   &H80000005&
      Height          =   210
      Left            =   6600
      TabIndex        =   3
      Top             =   1200
      Width           =   1245
   End
   Begin MSMask.MaskEdBox Beneficiario 
      Height          =   225
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   50
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   2235
      TabIndex        =   1
      Top             =   1245
      Width           =   1365
      _ExtentX        =   2408
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
   Begin MSMask.MaskEdBox Cheque 
      Height          =   225
      Left            =   1215
      TabIndex        =   0
      Top             =   1200
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridChequesPag3 
      Height          =   1860
      Left            =   135
      TabIndex        =   4
      Top             =   570
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   3281
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Qtde de Cheques:"
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
      Left            =   4560
      TabIndex        =   16
      Top             =   210
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Conta Corrente:"
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
      Left            =   825
      TabIndex        =   17
      Top             =   210
      Width           =   1350
   End
   Begin VB.Label LabelConta 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   180
      Width           =   1995
   End
   Begin VB.Label LabelQtdCheques 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6210
      TabIndex        =   19
      Top             =   180
      Width           =   1005
   End
End
Attribute VB_Name = "ChequesPag3Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTChequesPag3
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTChequesPag3
    Set objCT.objUserControl = Me
End Sub

Private Sub Atualizar_GotFocus()
     Call objCT.Atualizar_GotFocus
End Sub

Private Sub Atualizar_KeyPress(KeyAscii As Integer)
     Call objCT.Atualizar_KeyPress(KeyAscii)
End Sub

Private Sub Atualizar_Validate(Cancel As Boolean)
     Call objCT.Atualizar_Validate(Cancel)
End Sub

Private Sub BotaoConfigurarImpressao_Click()
     Call objCT.BotaoConfigurarImpressao_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoImprimirAPartir_Click()
     Call objCT.BotaoImprimirAPartir_Click
End Sub

Private Sub BotaoImprimirSelecao_Click()
     Call objCT.BotaoImprimirSelecao_Click
End Sub

Private Sub BotaoImprimirTeste_Click()
     Call objCT.BotaoImprimirTeste_Click
End Sub

Private Sub BotaoImprimirTudo_Click()
     Call objCT.BotaoImprimirTudo_Click
End Sub

Private Sub BotaoNumAuto_Click()
     Call objCT.BotaoNumAuto_Click
End Sub

Private Sub BotaoSeguir_Click()
     Call objCT.BotaoSeguir_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoVoltar_Click()
     Call objCT.BotaoVoltar_Click
End Sub

Private Sub Cheque_GotFocus()
     Call objCT.Cheque_GotFocus
End Sub

Private Sub Cheque_KeyPress(KeyAscii As Integer)
     Call objCT.Cheque_KeyPress(KeyAscii)
End Sub

Private Sub Cheque_Validate(Cancel As Boolean)
     Call objCT.Cheque_Validate(Cancel)
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

Private Sub Beneficiario_GotFocus()
     Call objCT.Beneficiario_GotFocus
End Sub

Private Sub Beneficiario_KeyPress(KeyAscii As Integer)
     Call objCT.Beneficiario_KeyPress(KeyAscii)
End Sub

Private Sub Beneficiario_Validate(Cancel As Boolean)
     Call objCT.Beneficiario_Validate(Cancel)
End Sub

Private Sub GridChequesPag3_Click()
     Call objCT.GridChequesPag3_Click
End Sub

Private Sub GridChequesPag3_GotFocus()
     Call objCT.GridChequesPag3_GotFocus
End Sub

Private Sub GridChequesPag3_EnterCell()
     Call objCT.GridChequesPag3_EnterCell
End Sub

Private Sub GridChequesPag3_LeaveCell()
     Call objCT.GridChequesPag3_LeaveCell
End Sub

Private Sub GridChequesPag3_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridChequesPag3_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridChequesPag3_KeyPress(KeyAscii As Integer)
     Call objCT.GridChequesPag3_KeyPress(KeyAscii)
End Sub

Private Sub GridChequesPag3_Validate(Cancel As Boolean)
     Call objCT.GridChequesPag3_Validate(Cancel)
End Sub

Private Sub GridChequesPag3_RowColChange()
     Call objCT.GridChequesPag3_RowColChange
End Sub

Private Sub GridChequesPag3_Scroll()
     Call objCT.GridChequesPag3_Scroll
End Sub

Function Trata_Parametros(Optional objChequesPag As ClassChequesPag) As Long
     Trata_Parametros = objCT.Trata_Parametros(objChequesPag)
End Function

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
Private Sub LabelConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelConta, Source, X, Y)
End Sub
Private Sub LabelConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelConta, Button, Shift, X, Y)
End Sub
Private Sub LabelQtdCheques_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelQtdCheques, Source, X, Y)
End Sub
Private Sub LabelQtdCheques_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelQtdCheques, Button, Shift, X, Y)
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

