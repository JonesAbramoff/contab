VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl EstoqueInicial2 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   9150
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6375
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   195
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "EstoqueInicial2.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "EstoqueInicial2.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "EstoqueInicial2.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "EstoqueInicial2.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox Padrao 
      Caption         =   "Padrão"
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
      Left            =   4740
      TabIndex        =   1
      Top             =   465
      Width           =   960
   End
   Begin VB.CheckBox Fixar 
      Caption         =   "Fixar"
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
      Left            =   4755
      TabIndex        =   3
      Top             =   960
      Width           =   810
   End
   Begin VB.ComboBox Almoxarifado 
      Height          =   315
      Left            =   1935
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   2520
   End
   Begin VB.TextBox LocalizacaoFisica 
      Height          =   300
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2565
      Width           =   3855
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   3015
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3630
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   315
      Left            =   1875
      TabIndex        =   6
      Top             =   3615
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   945
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView TvwProdutos 
      Height          =   3045
      Left            =   6195
      TabIndex        =   7
      Top             =   1110
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   5371
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSMask.MaskEdBox ContaContabil 
      Height          =   315
      Left            =   1875
      TabIndex        =   5
      Top             =   3105
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label ContaContabilLabel 
      AutoSize        =   -1  'True
      Caption         =   "Conta Estoque:"
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
      Left            =   465
      TabIndex        =   23
      Top             =   3135
      Width           =   1335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifado:"
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
      Left            =   690
      TabIndex        =   22
      Top             =   465
      Width           =   1170
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
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
      Left            =   1095
      TabIndex        =   21
      Top             =   945
      Width           =   750
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   900
      TabIndex        =   20
      Top             =   1500
      Width           =   945
   End
   Begin VB.Label DescricaoProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1935
      TabIndex        =   19
      Top             =   1470
      Width           =   3855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Unidade de Medida:"
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
      TabIndex        =   18
      Top             =   2070
      Width           =   1740
   End
   Begin VB.Label UnidMed 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1905
      TabIndex        =   17
      Top             =   2040
      Width           =   1650
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1305
      TabIndex        =   16
      Top             =   3660
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Localização Física:"
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
      Top             =   2610
      Width           =   1695
   End
   Begin VB.Label LabelProduto 
      AutoSize        =   -1  'True
      Caption         =   "Produtos"
      Height          =   195
      Left            =   6195
      TabIndex        =   14
      Top             =   885
      Width           =   645
   End
End
Attribute VB_Name = "EstoqueInicial2"
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
''    Parent.HelpContextID = IDH_ESTOQUE_INICIAL2
''    Set Form_Load_Ocx = Me
''    Caption = "Implantação - Produto X Almoxarifado"
''    Call Form_Load
''
''End Function
''
''Public Function Name() As String
''
''    Name = "EstoqueInicial2"
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
''

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub DescricaoProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoProduto, Source, X, Y)
End Sub

Private Sub DescricaoProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoProduto, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub UnidMed_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidMed, Source, X, Y)
End Sub

Private Sub UnidMed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidMed, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

