VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ConsumoRecalculo 
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   5580
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   2820
      ScaleHeight     =   690
      ScaleWidth      =   2445
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   195
      Width           =   2505
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   1935
         Picture         =   "ConsumoRecalculo.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   405
      End
      Begin VB.CommandButton BotaoApurar 
         Height          =   510
         Left            =   120
         Picture         =   "ConsumoRecalculo.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   510
         Left            =   1455
         Picture         =   "ConsumoRecalculo.ctx":1A40
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   390
      End
   End
   Begin VB.ComboBox Ano 
      Height          =   315
      Left            =   825
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   420
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      Height          =   870
      Left            =   165
      TabIndex        =   12
      Top             =   1110
      Width           =   5175
      Begin VB.ComboBox PeriodoDe 
         Height          =   315
         Left            =   735
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   375
         Width           =   1680
      End
      Begin VB.ComboBox PeriodoAte 
         Height          =   315
         Left            =   3075
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   390
         Width           =   1680
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
         TabIndex        =   14
         Top             =   435
         Width           =   315
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
         Left            =   2670
         TabIndex        =   13
         Top             =   435
         Width           =   360
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Produtos"
      Height          =   810
      Left            =   150
      TabIndex        =   9
      Top             =   2160
      Width           =   5205
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   300
         Left            =   735
         TabIndex        =   3
         Top             =   330
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoAte 
         Height          =   300
         Left            =   3075
         TabIndex        =   4
         Top             =   330
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProdutoDe 
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
         Height          =   195
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   375
         Width           =   315
      End
      Begin VB.Label LabelProdutoAte 
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
         Left            =   2670
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   375
         Width           =   360
      End
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
      Left            =   360
      TabIndex        =   15
      Top             =   450
      Width           =   405
   End
End
Attribute VB_Name = "ConsumoRecalculo"
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
''    Parent.HelpContextID = IDH_CONSUMO_RECALCULO
''    Set Form_Load_Ocx = Me
''    Caption = "Recálculo do Consumo"
''    Call Form_Load
''
''End Function
''
''Public Function Name() As String
''
''    Name = "ConsumoRecalculo"
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

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

