VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl SerieNFiscal 
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   ScaleHeight     =   2940
   ScaleWidth      =   5355
   Begin VB.CheckBox Eletronica 
      Caption         =   "Eletrônica"
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
      Left            =   3915
      TabIndex        =   3
      Top             =   915
      Width           =   1290
   End
   Begin VB.TextBox NomeTsk 
      Height          =   315
      Left            =   1770
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1920
      Width           =   3195
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
      Left            =   1890
      TabIndex        =   1
      Top             =   390
      Width           =   945
   End
   Begin VB.ComboBox TipoFormulario 
      Height          =   315
      ItemData        =   "SerieNFiscalEver.ctx":0000
      Left            =   1770
      List            =   "SerieNFiscalEver.ctx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1410
      Width           =   3195
   End
   Begin VB.ComboBox Serie 
      Height          =   315
      Left            =   900
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   330
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3150
      ScaleHeight     =   495
      ScaleWidth      =   2025
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "SerieNFiscalEver.ctx":0099
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   555
         Picture         =   "SerieNFiscalEver.ctx":01F3
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1050
         Picture         =   "SerieNFiscalEver.ctx":037D
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "SerieNFiscalEver.ctx":08AF
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox ProxNumNFiscal 
      Height          =   300
      Left            =   3120
      TabIndex        =   2
      Top             =   877
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaxLinhasNF 
      Height          =   300
      Left            =   1770
      TabIndex        =   6
      Top             =   2430
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   " "
   End
   Begin VB.Label LabelMaxLinhas 
      AutoSize        =   -1  'True
      Caption         =   "Máximo de Linhas:"
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
      Left            =   135
      TabIndex        =   16
      Top             =   2475
      Width           =   1590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "NomeTsk:"
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
      Left            =   840
      TabIndex        =   15
      Top             =   1980
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Formulário:"
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
      Left            =   330
      TabIndex        =   12
      Top             =   1470
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Próximo Número de Nota Fiscal:"
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
      Left            =   330
      TabIndex        =   13
      Top             =   930
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Série:"
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
      Left            =   330
      TabIndex        =   14
      Top             =   390
      Width           =   510
   End
End
Attribute VB_Name = "SerieNFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTSerieNFiscal
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTSerieNFiscal
    Set objCT.objUserControl = Me
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Function Trata_Parametros(Optional objSerie As ClassSerie) As Long
     Trata_Parametros = objCT.Trata_Parametros(objSerie)
End Function

Private Sub Padrao_Click()
     Call objCT.Padrao_Click
End Sub

Private Sub ProxNumNFiscal_Change()
     Call objCT.ProxNumNFiscal_Change
End Sub

Private Sub ProxNumNFiscal_GotFocus()
     Call objCT.ProxNumNFiscal_GotFocus
End Sub

Private Sub ProxNumNFiscal_Validate(Cancel As Boolean)
     Call objCT.ProxNumNFiscal_Validate(Cancel)
End Sub

Private Sub Serie_Change()
     Call objCT.Serie_Change
End Sub

Private Sub Serie_Click()
     Call objCT.Serie_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub Serie_Remove(objSerie As ClassSerie)
     Call objCT.Serie_Remove(objSerie)
End Sub

Private Sub Serie_Adiciona(objSerie As ClassSerie)
     Call objCT.Serie_Adiciona(objSerie)
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Serie_Validate(Cancel As Boolean)
     Call objCT.Serie_Validate(Cancel)
End Sub

Private Sub TipoFormulario_Change()
     Call objCT.TipoFormulario_Change
End Sub

Private Sub TipoFormulario_Click()
     Call objCT.TipoFormulario_Click
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub
Private Sub MaxLinhasNF_Change()
     Call objCT.MaxLinhasNF_Change
End Sub

Private Sub MaxLinhasNF_GotFocus()
     Call objCT.MaxLinhasNF_GotFocus
End Sub

Private Sub MaxLinhasNF_Validate(Cancel As Boolean)
     Call objCT.MaxLinhasNF_Validate(Cancel)
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

'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call objCT.UserControl_KeyDown(KeyCode, Shift)
'End Sub

