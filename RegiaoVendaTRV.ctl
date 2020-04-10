VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RegiaoVenda 
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   7245
   Begin VB.ComboBox UsuRespCallCenter 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2460
      Width           =   3000
   End
   Begin VB.ComboBox Pais 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "RegiaoVendaTRV.ctx":0000
      Left            =   1680
      List            =   "RegiaoVendaTRV.ctx":0002
      TabIndex        =   3
      Top             =   1155
      Width           =   2490
   End
   Begin VB.ListBox Regioes 
      Height          =   2205
      ItemData        =   "RegiaoVendaTRV.ctx":0004
      Left            =   4875
      List            =   "RegiaoVendaTRV.ctx":0006
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   1005
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4875
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RegiaoVendaTRV.ctx":0008
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RegiaoVendaTRV.ctx":0186
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RegiaoVendaTRV.ctx":06B8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RegiaoVendaTRV.ctx":0842
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2220
      Picture         =   "RegiaoVendaTRV.ctx":099C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   345
      Width           =   300
   End
   Begin VB.ComboBox ComboCobrador 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1995
      Width           =   2985
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   330
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   735
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Gerente 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   30
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CodRegiaoPai 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   2895
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      Caption         =   "Resp. Call Center:"
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
      TabIndex        =   22
      Top             =   2520
      Width           =   1560
   End
   Begin VB.Label DescRegiaoPai 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2235
      TabIndex        =   21
      Top             =   2880
      Width           =   2445
   End
   Begin VB.Label LabelRegiaoPai 
      AutoSize        =   -1  'True
      Caption         =   "Região Pai:"
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
      Left            =   600
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   20
      Top             =   2955
      Width           =   1005
   End
   Begin VB.Label Label70 
      AutoSize        =   -1  'True
      Caption         =   "Gerente:"
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
      Left            =   840
      TabIndex        =   19
      Top             =   1620
      Width           =   750
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "País:"
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
      TabIndex        =   18
      Top             =   1215
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Regiões de Venda"
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
      Left            =   4875
      TabIndex        =   17
      Top             =   780
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   930
      TabIndex        =   16
      Top             =   375
      Width           =   660
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   660
      TabIndex        =   15
      Top             =   780
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Usuário Cobrador:"
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
      Left            =   105
      TabIndex        =   14
      Top             =   2055
      Width           =   1545
   End
End
Attribute VB_Name = "RegiaoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTRegiaoVenda
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoProxNum_Click()
    Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoExcluir_Click()
    Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
    Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
    Call objCT.BotaoGravar_Click
End Sub

Private Sub Codigo_GotFocus()
    Call objCT.Codigo_GotFocus
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
    Call objCT.Codigo_Validate(Cancel)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

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

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub Descricao_Change()
     Call objCT.Descricao_Change
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros(Optional objRegiaoVenda As ClassRegiaoVenda) As Long
     Trata_Parametros = objCT.Trata_Parametros(objRegiaoVenda)
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Gerente_Change()
    Call objCT.Gerente_Change
End Sub

Private Sub Pais_Change()
    Call objCT.Pais_Change
End Sub

Private Sub Pais_Click()
    Call objCT.Pais_Click
End Sub

Private Sub Pais_Validate(Cancel As Boolean)
    Call objCT.Pais_Validate(Cancel)
End Sub

Private Sub Regioes_DblClick()
    Call objCT.Regioes_DblClick
End Sub


Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

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

Private Sub UserControl_Initialize()

    Set objCT = New CTRegiaoVenda
    Set objCT.objUserControl = Me
    
    Set objCT.gobjInfoUsu = New CTRegiaoVendaVGTRV
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTRegiaoVendaTRV

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
    Call objCT.gobjInfoUsu.gobjTelaUsu.UserControl_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub ComboCobrador_Click()
    Call objCT.ComboCobrador_Click
End Sub

Private Sub ComboCobrador_Validate(Cancel As Boolean)
    Call objCT.ComboCobrador_Validate(Cancel)
End Sub

Private Sub CodRegiaoPai_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CodRegiaoPai_Change(objCT)
End Sub

Private Sub CodRegiaoPai_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CodRegiaoPai_GotFocus(objCT)
End Sub

Private Sub CodRegiaoPai_KeyPress(KeyAscii As Integer)
'     Call objCT.gobjInfoUsu.gobjTelaUsu.CodRegiaoPai_KeyPress(objCT, KeyAscii)
End Sub

Private Sub CodRegiaoPai_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.CodRegiaoPai_Validate(objCT, Cancel)
End Sub

Private Sub LabelRegiaoPai_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.LabelRegiaoPai_Click(objCT)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label63_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label63, Source, X, Y)
End Sub

Private Sub Label63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label63, Button, Shift, X, Y)
End Sub

Private Sub Label70_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label70, Source, X, Y)
End Sub

Private Sub Label70_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label70, Button, Shift, X, Y)
End Sub

Private Sub UsuRespCallCenter_Click()
    objCT.UsuRespCallCenter_Click
End Sub

Private Sub UsuRespCallCenter_Validate(Cancel As Boolean)
    objCT.UsuRespCallCenter_Validate (Cancel)
End Sub
