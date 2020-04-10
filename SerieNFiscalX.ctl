VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl SerieNFiscal 
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   ScaleHeight     =   3720
   ScaleWidth      =   5940
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3705
      ScaleHeight     =   495
      ScaleWidth      =   2025
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   135
      Width           =   2085
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "SerieNFiscalX.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1050
         Picture         =   "SerieNFiscalX.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   555
         Picture         =   "SerieNFiscalX.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "SerieNFiscalX.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressão"
      Height          =   645
      Left            =   105
      TabIndex        =   10
      Top             =   2625
      Width           =   5625
      Begin VB.TextBox NomeTsk 
         Height          =   315
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   12
         Top             =   210
         Width           =   3195
      End
      Begin VB.ComboBox TipoFormulario 
         Height          =   315
         ItemData        =   "SerieNFiscalX.ctx":0994
         Left            =   1455
         List            =   "SerieNFiscalX.ctx":09AA
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   4020
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
         Left            =   105
         TabIndex        =   14
         Top             =   270
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
         Left            =   45
         TabIndex        =   13
         Top             =   660
         Visible         =   0   'False
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Próximo número de Nota Fiscal"
      Height          =   675
      Left            =   90
      TabIndex        =   7
      Top             =   1890
      Width           =   5640
      Begin MSMask.MaskEdBox ProxNumNFiscal 
         Height          =   300
         Left            =   1020
         TabIndex        =   8
         Top             =   270
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   285
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificação"
      Height          =   1110
      Left            =   75
      TabIndex        =   0
      Top             =   705
      Width           =   5670
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
         Left            =   1950
         TabIndex        =   4
         Top             =   345
         Width           =   945
      End
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   1050
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   285
         Width           =   855
      End
      Begin VB.CheckBox Eletronica 
         Caption         =   "Eletrônica Federal"
         Enabled         =   0   'False
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
         Left            =   3630
         TabIndex        =   2
         Top             =   330
         Width           =   1980
      End
      Begin VB.ComboBox ModDocFis 
         Height          =   315
         ItemData        =   "SerieNFiscalX.ctx":0A2D
         Left            =   1050
         List            =   "SerieNFiscalX.ctx":0A43
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   675
         Width           =   4560
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
         Left            =   495
         TabIndex        =   6
         Top             =   345
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   315
         TabIndex        =   5
         Top             =   735
         Width           =   690
      End
   End
   Begin MSMask.MaskEdBox MaxLinhasNF 
      Height          =   300
      Left            =   1755
      TabIndex        =   20
      Top             =   3345
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
      Left            =   120
      TabIndex        =   21
      Top             =   3390
      Width           =   1590
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

Private Sub ModDocFis_Change()
     Call objCT.ModDocFis_Change
End Sub

Private Sub ModDocFis_Click()
     Call objCT.ModDocFis_Click
End Sub

