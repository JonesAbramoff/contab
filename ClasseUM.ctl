VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ClasseUM 
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   6990
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4575
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   225
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ClasseUM.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ClasseUM.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ClasseUM.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ClasseUM.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unidade Base"
      Height          =   810
      Left            =   150
      TabIndex        =   8
      Top             =   1560
      Width           =   5205
      Begin MSMask.MaskEdBox SiglaUMBase 
         Height          =   315
         Left            =   855
         TabIndex        =   2
         Top             =   315
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeUMBase 
         Height          =   315
         Left            =   2655
         TabIndex        =   3
         Top             =   330
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sigla:"
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
         Left            =   300
         TabIndex        =   17
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Left            =   2025
         TabIndex        =   16
         Top             =   390
         Width           =   555
      End
   End
   Begin VB.TextBox Conversao 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   5235
      TabIndex        =   7
      Top             =   3015
      Width           =   330
   End
   Begin MSMask.MaskEdBox Nome 
      Height          =   225
      Left            =   1860
      TabIndex        =   5
      Top             =   3015
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Sigla 
      Height          =   225
      Left            =   855
      TabIndex        =   4
      Top             =   3030
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   5
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   225
      Left            =   4110
      TabIndex        =   6
      Top             =   3015
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
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
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SiglaUMBase1 
      Height          =   225
      Left            =   5595
      TabIndex        =   9
      Top             =   3030
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1185
      TabIndex        =   1
      Top             =   1020
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridUM 
      Height          =   2070
      Left            =   150
      TabIndex        =   10
      Top             =   2970
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   3651
      _Version        =   393216
      Rows            =   8
      Cols            =   5
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1185
      TabIndex        =   0
      Top             =   450
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   450
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   21
      Top             =   510
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
      Left            =   180
      TabIndex        =   20
      Top             =   1050
      Width           =   930
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Unidades de Medida"
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
      Left            =   195
      TabIndex        =   19
      Top             =   2745
      Width           =   1755
   End
   Begin VB.Label LabelConversao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conversão"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   2715
      Width           =   2535
   End
End
Attribute VB_Name = "ClasseUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'responsavel Jones
'revisada em 29/08/98
'pendencias:
' ver observacoes em MATGrava

Option Explicit

Event Unload()

Private WithEvents objCT As CTClasseUM
Attribute objCT.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me
    
End Function

Function Trata_Parametros(Optional objClasseUM As ClassClasseUM) As Long
'Se a classeUM vier preenchida colocá-la na tela

    Trata_Parametros = objCT.Trata_Parametros(objClasseUM)

End Function

Private Sub BotaoExcluir_Click()
'Chama ClasseUM_Excluir

    Call objCT.BotaoExcluir_Click

End Sub

Private Sub Codigo_Change()

    Call objCT.Codigo_Change

End Sub

Private Sub Codigo_GotFocus()
    
    Call objCT.Codigo_GotFocus

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

    Call objCT.Codigo_Validate(Cancel)
    
End Sub

Private Sub Descricao_Change()

    Call objCT.Descricao_Change

End Sub

Public Sub Form_Activate()

    Call objCT.Form_Activate

End Sub

Public Sub Form_Deactivate()

    Call objCT.Form_Deactivate

End Sub

Private Sub BotaoFechar_Click()

    Call objCT.BotaoFechar_Click

End Sub

Private Sub BotaoGravar_Click()

    Call objCT.BotaoGravar_Click

End Sub

Private Sub BotaoLimpar_Click()
'Chama a função que limpa toda a tela

    Call objCT.BotaoLimpar_Click

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
        If Cancel = False Then
            Set objCT.objUserControl = Nothing
            Set objCT = Nothing
        End If
    End If
    
End Sub

Private Sub GridUM_Click()

    Call objCT.GridUM_Click

End Sub

Private Sub GridUM_GotFocus()

    Call objCT.GridUM_GotFocus

End Sub

Private Sub GridUM_EnterCell()

    Call objCT.GridUM_EnterCell

End Sub

Private Sub GridUM_LeaveCell()

    Call objCT.GridUM_LeaveCell

End Sub

Private Sub GridUM_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objCT.GridUM_KeyDown(KeyCode, Shift)

End Sub

Private Sub GridUM_KeyPress(KeyAscii As Integer)

    Call objCT.GridUM_KeyPress(KeyAscii)

End Sub

Private Sub GridUM_Validate(Cancel As Boolean)

    Call objCT.GridUM_Validate(Cancel)

End Sub

Private Sub GridUM_RowColChange()

    Call objCT.GridUM_RowColChange

End Sub

Private Sub GridUM_Scroll()

    Call objCT.GridUM_Scroll

End Sub

Private Sub LabelCodigo_Click()

    Call objCT.LabelCodigo_Click

End Sub

Private Sub NomeUMBase_Change()

        Call objCT.NomeUMBase_Change

End Sub

Private Sub objCT_Unload()

    RaiseEvent Unload

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)



End Sub

Private Sub SiglaUMBase_Change()

        Call objCT.SiglaUMBase_Change

End Sub

Private Sub SiglaUMBase_Validate(Cancel As Boolean)

    Call objCT.SiglaUMBase_Validate(Cancel)

End Sub


Private Sub Quantidade_Change()

    Call objCT.Quantidade_Change

End Sub

Private Sub Quantidade_GotFocus()

    Call objCT.Quantidade_GotFocus

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call objCT.Quantidade_KeyPress(KeyAscii)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

    Call objCT.Quantidade_Validate(Cancel)

End Sub

Private Sub Nome_Change()

    Call objCT.Nome_Change

End Sub

Private Sub Nome_GotFocus()

    Call objCT.Nome_GotFocus

End Sub

Private Sub Nome_KeyPress(KeyAscii As Integer)

    Call objCT.Nome_KeyPress(KeyAscii)

End Sub

Private Sub Nome_Validate(Cancel As Boolean)

    Call objCT.Nome_Validate(Cancel)

End Sub

Private Sub Sigla_Change()

    Call objCT.Sigla_Change

End Sub

Private Sub Sigla_GotFocus()

    Call objCT.Sigla_GotFocus

End Sub

Private Sub Sigla_KeyPress(KeyAscii As Integer)

    Call objCT.Sigla_KeyPress(KeyAscii)

End Sub

Private Sub Sigla_Validate(Cancel As Boolean)

    Call objCT.Sigla_Validate(Cancel)

End Sub

Private Sub SiglaUMBase1_Change()
    
    Call objCT.SiglaUMBase1_Change

End Sub

Private Sub SiglaUMBase1_GotFocus()

    Call objCT.SiglaUMBase1_GotFocus

End Sub

Private Sub SiglaUMBase1_KeyPress(KeyAscii As Integer)

    Call objCT.SiglaUMBase1_KeyPress(KeyAscii)

End Sub

Private Sub SiglaUMBase1_Validate(Cancel As Boolean)

    Call objCT.SiglaUMBase1_Validate(Cancel)

End Sub

Private Sub Conversao_Change()
    
    Call objCT.Conversao_Change

End Sub

Private Sub Conversao_GotFocus()

    Call objCT.Conversao_GotFocus

End Sub

Private Sub Conversao_KeyPress(KeyAscii As Integer)

    Call objCT.Conversao_KeyPress(KeyAscii)

End Sub

Private Sub Conversao_Validate(Cancel As Boolean)

    Call objCT.Conversao_Validate(Cancel)

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

Private Sub UserControl_Initialize()
    Set objCT = New CTClasseUM
    Set objCT.objUserControl = Me
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'Private Sub Unload(objme As Object)
'
'   RaiseEvent Unload
'
'End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call objCT.UserControl_KeyDown(KeyCode, Shift)

End Sub


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

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub LabelConversao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelConversao, Source, X, Y)
End Sub

Private Sub LabelConversao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelConversao, Button, Shift, X, Y)
End Sub

