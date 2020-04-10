VERSION 5.00
Begin VB.UserControl ConsumoMedio 
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   LockControls    =   -1  'True
   ScaleHeight     =   1125
   ScaleWidth      =   6240
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   3465
      ScaleHeight     =   690
      ScaleWidth      =   2445
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Width           =   2505
      Begin VB.CommandButton BotaoLimpar 
         Height          =   510
         Left            =   1455
         Picture         =   "ConsumoMedio.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   90
         Width           =   390
      End
      Begin VB.CommandButton BotaoApurar 
         Height          =   510
         Left            =   120
         Picture         =   "ConsumoMedio.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   1935
         Picture         =   "ConsumoMedio.ctx":1DF4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
   End
   Begin VB.ComboBox Ano 
      Height          =   315
      Left            =   615
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
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
      Left            =   150
      TabIndex        =   5
      Top             =   390
      Width           =   405
   End
End
Attribute VB_Name = "ConsumoMedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''Option Explicit
''
'''Property Variables:
''Dim m_Caption As String
''Event Unload()
''
''
''
''
'''**** inicio do trecho a ser copiado *****
''
''Public Function Form_Load_Ocx() As Object
''    Parent.HelpContextID = IDH_CONSUMO_MEDIO
''    Set Form_Load_Ocx = Me
''    Caption = "Cálculo do Consumo Médio"
''    Call Form_Load
''
''End Function
''
''Public Function Name() As String
''
''    Name = "ConsumoMedio"
''
''End Function
''
''Public Sub Show()
''    Parent.Show
''    Parent.SetFocus
''End Sub
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=UserControl,UserControl,-1,Controls
''Public Property Get Controls() As Object
''    Set Controls = UserControl.Controls
''End Property
''
''Public Property Get hWnd() As Long
''    hWnd = UserControl.hWnd
''End Property
''
''Public Property Get Height() As Long
''    Height = UserControl.Height
''End Property
''
''Public Property Get Width() As Long
''    Width = UserControl.Width
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=UserControl,UserControl,-1,ActiveControl
''Public Property Get ActiveControl() As Object
''    Set ActiveControl = UserControl.ActiveControl
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=UserControl,UserControl,-1,Enabled
''Public Property Get Enabled() As Boolean
''    Enabled = UserControl.Enabled
''End Property
''
''Public Property Let Enabled(ByVal New_Enabled As Boolean)
''    UserControl.Enabled() = New_Enabled
''    PropertyChanged "Enabled"
''End Property
''
'''Load property values from storage
''Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
''
''    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
''End Sub
''
'''Write property values to storage
''Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
''
''    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
''End Sub
''
''Private Sub Unload(objme As Object)
''
''   RaiseEvent Unload
''
''End Sub
''
''Public Property Get Caption() As String
''    Caption = m_Caption
''End Property
''
''Public Property Let Caption(ByVal New_Caption As String)
''    Parent.Caption = New_Caption
''    m_Caption = New_Caption
''End Property
''
'''**** fim do trecho a ser copiado *****
''

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

