VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl EstoqueInicial1 
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   8670
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6315
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   195
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "EstoqueInicial1.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "EstoqueInicial1.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "EstoqueInicial1.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "EstoqueInicial1.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
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
      Left            =   4695
      TabIndex        =   1
      Top             =   450
      Width           =   945
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
      Left            =   4710
      TabIndex        =   3
      Top             =   945
      Width           =   795
   End
   Begin VB.ComboBox Almoxarifado 
      Height          =   315
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   390
      Width           =   2505
   End
   Begin VB.TextBox LocalizacaoFisica 
      Height          =   300
      Left            =   1845
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2550
      Width           =   3840
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   315
      Left            =   1860
      TabIndex        =   6
      Top             =   3660
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Left            =   1815
      TabIndex        =   7
      Top             =   4200
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
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
   Begin MSMask.MaskEdBox Produto 
      Height          =   285
      Left            =   1875
      TabIndex        =   2
      Top             =   930
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView TvwProdutos 
      Height          =   3195
      Left            =   6150
      TabIndex        =   8
      Top             =   1095
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   5636
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSMask.MaskEdBox ContaContabil 
      Height          =   315
      Left            =   1830
      TabIndex        =   5
      Top             =   3090
      Width           =   1635
      _ExtentX        =   2884
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
      Left            =   420
      TabIndex        =   24
      Top             =   3120
      Width           =   1320
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
      Left            =   645
      TabIndex        =   23
      Top             =   450
      Width           =   1155
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
      Left            =   1050
      TabIndex        =   22
      Top             =   930
      Width           =   735
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
      Left            =   855
      TabIndex        =   21
      Top             =   1485
      Width           =   930
   End
   Begin VB.Label DescricaoProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1890
      TabIndex        =   20
      Top             =   1455
      Width           =   3840
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
      Left            =   75
      TabIndex        =   19
      Top             =   2055
      Width           =   1725
   End
   Begin VB.Label UnidMed 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1860
      TabIndex        =   18
      Top             =   2025
      Width           =   1635
   End
   Begin VB.Label Label5 
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
      Height          =   240
      Left            =   1200
      TabIndex        =   17
      Top             =   4245
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
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
      Left            =   720
      TabIndex        =   16
      Top             =   3705
      Width           =   1050
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
      Left            =   120
      TabIndex        =   15
      Top             =   2595
      Width           =   1680
   End
   Begin VB.Label LabelProduto 
      AutoSize        =   -1  'True
      Caption         =   "Produtos"
      Height          =   195
      Left            =   6150
      TabIndex        =   14
      Top             =   870
      Width           =   630
   End
End
Attribute VB_Name = "EstoqueInicial1"
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
''    Parent.HelpContextID = IDH_ESTOQUE_INICIAL1
''    Set Form_Load_Ocx = Me
''    Caption = "Implantação - Estoques Iniciais"
''    Call Form_Load
''
''End Function
''
''Public Function Name() As String
''
''    Name = "EstoqueInicial1"
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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

