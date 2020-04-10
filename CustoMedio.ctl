VERSION 5.00
Begin VB.UserControl CustoMedio 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   LockControls    =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   5880
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   3150
      ScaleHeight     =   690
      ScaleWidth      =   2445
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   225
      Width           =   2505
      Begin VB.CommandButton BotaoLimpar 
         Height          =   510
         Left            =   1455
         Picture         =   "CustoMedio.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   390
      End
      Begin VB.CommandButton BotaoApurar 
         Height          =   510
         Left            =   120
         Picture         =   "CustoMedio.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   1935
         Picture         =   "CustoMedio.ctx":1DF4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Apropriação a ser utilizada no Recálculo"
      Height          =   885
      Left            =   165
      TabIndex        =   12
      Top             =   2145
      Width           =   5520
      Begin VB.OptionButton Sequencial 
         Caption         =   "Sequencial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   510
         TabIndex        =   1
         Top             =   435
         Width           =   1275
      End
      Begin VB.OptionButton Diaria 
         Caption         =   "Diária"
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
         Left            =   2535
         TabIndex        =   2
         Top             =   435
         Width           =   870
      End
      Begin VB.OptionButton Mensal 
         Caption         =   "Mensal"
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
         Left            =   4155
         TabIndex        =   3
         Top             =   435
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período"
      Height          =   870
      Left            =   195
      TabIndex        =   8
      Top             =   1110
      Width           =   5475
      Begin VB.ComboBox PeriodoAte 
         Height          =   315
         Left            =   3480
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   390
         Width           =   1680
      End
      Begin VB.Label Label4 
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
         Height          =   195
         Left            =   3075
         TabIndex        =   11
         Top             =   435
         Width           =   360
      End
      Begin VB.Label Label3 
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
         Left            =   300
         TabIndex        =   10
         Top             =   435
         Width           =   315
      End
      Begin VB.Label PeriodoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   750
         TabIndex        =   9
         Top             =   390
         Width           =   1680
      End
   End
End
Attribute VB_Name = "CustoMedio"
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
''    Parent.HelpContextID = IDH_CUSTO_MEDIO
''    Set Form_Load_Ocx = Me
''    Caption = "Recálculo do Custo Médio"
''    Call Form_Load
''
''End Function
''
''Public Function Name() As String
''
''    Name = "CustoMedio"
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
''


Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub PeriodoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PeriodoDe, Source, X, Y)
End Sub

Private Sub PeriodoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PeriodoDe, Button, Shift, X, Y)
End Sub

