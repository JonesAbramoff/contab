VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ChequePagAvulso3 
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   ScaleHeight     =   3450
   ScaleWidth      =   6315
   Begin VB.CommandButton BotaoReter 
      Caption         =   "Reter Nominal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4290
      Picture         =   "ChequePagAvulso3Hic.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1995
      Width           =   1920
   End
   Begin VB.TextBox Nominal 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1140
      Width           =   3990
   End
   Begin VB.TextBox Observacao 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   1530
      Width           =   3990
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   1815
      ScaleHeight     =   495
      ScaleWidth      =   2625
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2790
      Width           =   2685
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   90
         Picture         =   "ChequePagAvulso3Hic.ctx":0432
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2115
         Picture         =   "ChequePagAvulso3Hic.ctx":0B90
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1102
         Picture         =   "ChequePagAvulso3Hic.ctx":0D0E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.CommandButton ConfigurarImpressao 
      Caption         =   "Configurar Impressão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      Picture         =   "ChequePagAvulso3Hic.ctx":14A0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1980
      Width           =   1920
   End
   Begin MSMask.MaskEdBox NumCheque 
      Height          =   300
      Left            =   1575
      TabIndex        =   0
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Imprimir 
      Caption         =   "Imprimir o Cheque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2205
      Picture         =   "ChequePagAvulso3Hic.ctx":2042
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1980
      Width           =   1920
   End
   Begin VB.Label Label3 
      Caption         =   "Observação:"
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
      Left            =   405
      TabIndex        =   15
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label LabelConta 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3420
      TabIndex        =   9
      Top             =   150
      Width           =   2085
   End
   Begin VB.Label Label1 
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
      Height          =   195
      Index           =   0
      Left            =   2745
      TabIndex        =   10
      Top             =   180
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "No. do Cheque:"
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
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   11
      Top             =   165
      Width           =   1395
   End
   Begin VB.Label Label2 
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
      Height          =   180
      Left            =   990
      TabIndex        =   12
      Top             =   765
      Width           =   555
   End
   Begin VB.Label LabelValorCheque 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1575
      TabIndex        =   13
      Top             =   720
      Width           =   1860
   End
   Begin VB.Label Label4 
      Caption         =   "Nominal à:"
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
      Left            =   600
      TabIndex        =   14
      Top             =   1230
      Width           =   945
   End
End
Attribute VB_Name = "ChequePagAvulso3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTChequePagAvulso3
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTChequePagAvulso3
    Set objCT.objUserControl = Me

    '################################################
    'Inserido por Wagner18/04/2005
    'hicare
    Set objCT.gobjInfoUsu = New CTChequesPag3AVGHic
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTChequesPag3AHic
    '################################################
    
End Sub

Function Trata_Parametros(Optional objChequesPagAvulso As ClassChequesPagAvulso) As Long
     Trata_Parametros = objCT.Trata_Parametros(objChequesPagAvulso)
End Function

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoSeguir_Click()
     Call objCT.BotaoSeguir_Click
End Sub

Private Sub BotaoVoltar_Click()
     Call objCT.BotaoVoltar_Click
End Sub

Private Sub ConfigurarImpressao_Click()
     Call objCT.ConfigurarImpressao_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub Imprimir_Click()
     Call objCT.Imprimir_Click
End Sub

Private Sub NumCheque_GotFocus()
     Call objCT.NumCheque_GotFocus
End Sub

Private Sub NumCheque_Validate(bCancel As Boolean)
     Call objCT.NumCheque_Validate(bCancel)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(label1(Index), Button, Shift, X, Y)
End Sub
Private Sub LabelConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelConta, Source, X, Y)
End Sub
Private Sub LabelConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelConta, Button, Shift, X, Y)
End Sub
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Private Sub LabelValorCheque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValorCheque, Source, X, Y)
End Sub
Private Sub LabelValorCheque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValorCheque, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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
    Call objCT.Name
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

'#########################################
'Inserido por Wagner
Private Sub BotaoReter_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoReter_Click(objCT)
End Sub
'#########################################

