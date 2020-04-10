VERSION 5.00
Begin VB.UserControl Anotacoes 
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   7800
   Begin VB.Frame FrameID 
      Caption         =   "Identificação"
      Height          =   1230
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   5145
      Begin VB.Label IdOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   945
         TabIndex        =   13
         Top             =   720
         Width           =   3885
      End
      Begin VB.Label LabelID 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
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
         Left            =   660
         TabIndex        =   12
         Top             =   788
         Width           =   270
      End
      Begin VB.Label Origem 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   945
         TabIndex        =   11
         Top             =   315
         Width           =   3885
      End
      Begin VB.Label LabelOrigem 
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
         Left            =   270
         TabIndex        =   10
         Top             =   383
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   600
      Left            =   5415
      ScaleHeight     =   540
      ScaleWidth      =   2190
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   165
      Width           =   2250
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1680
         Picture         =   "AnotacoesArtmill.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1155
         Picture         =   "AnotacoesArtmill.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   630
         Picture         =   "AnotacoesArtmill.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "AnotacoesArtmill.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   105
         Width           =   420
      End
   End
   Begin VB.TextBox Anotacao 
      Height          =   4515
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1905
      Width           =   7485
   End
   Begin VB.TextBox Titulo 
      Height          =   285
      Left            =   1050
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.Label Data 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6300
      TabIndex        =   15
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label LabelData 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
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
      Left            =   5730
      TabIndex        =   14
      Top             =   1515
      Width           =   480
   End
   Begin VB.Label LabelTexto 
      AutoSize        =   -1  'True
      Caption         =   "Texto:"
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
      Left            =   195
      TabIndex        =   8
      Top             =   1530
      Width           =   555
   End
   Begin VB.Label LabelTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Título:"
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
      Left            =   450
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   1530
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "Anotacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTAnotacoes
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTAnotacoes
    Set objCT.objUserControl = Me
End Sub

Public Function Trata_Parametros(ByVal objAnotacoes As ClassAnotacoes) As Long
     Trata_Parametros = objCT.Trata_Parametros(objAnotacoes)
End Function

Private Sub Anotacao_Change()
     Call objCT.Anotacao_Change
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub LabelTitulo_Click()
     Call objCT.LabelTitulo_Click
End Sub

Private Sub Titulo_Change()
     Call objCT.Titulo_Change
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub LabelData_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub
Private Sub LabelData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub
Private Sub LabelTitulo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTitulo, Source, X, Y)
End Sub
Private Sub LabelTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTitulo, Button, Shift, X, Y)
End Sub
Private Sub LabelOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrigem, Source, X, Y)
End Sub
Private Sub LabelOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrigem, Button, Shift, X, Y)
End Sub
Private Sub LabelID_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelID, Source, X, Y)
End Sub
Private Sub LabelID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelID, Button, Shift, X, Y)
End Sub
Private Sub LabelTexto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTexto, Source, X, Y)
End Sub
Private Sub LabelTexto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTexto, Button, Shift, X, Y)
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

