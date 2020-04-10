VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BorderoCobranca 
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   ScaleHeight     =   3960
   ScaleWidth      =   7845
   Begin VB.Frame Frame1 
      Caption         =   "Filtro dos Títulos"
      Height          =   2070
      Left            =   150
      TabIndex        =   5
      Top             =   1125
      Width           =   7545
      Begin VB.ComboBox Estado 
         Height          =   315
         Left            =   5220
         TabIndex        =   11
         Top             =   360
         Width           =   840
      End
      Begin VB.ComboBox TipoCliente 
         Height          =   315
         Left            =   5205
         TabIndex        =   10
         Top             =   804
         Width           =   2016
      End
      Begin VB.Frame Frame2 
         Caption         =   "Aceite"
         Height          =   630
         Left            =   180
         TabIndex        =   6
         Top             =   1230
         Width           =   7290
         Begin VB.OptionButton OptAceite 
            Caption         =   "Ambos"
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
            Index           =   0
            Left            =   255
            TabIndex        =   9
            Top             =   285
            Width           =   1500
         End
         Begin VB.OptionButton OptAceite 
            Caption         =   "Sem Aceite"
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
            Index           =   1
            Left            =   2617
            TabIndex        =   8
            Top             =   285
            Width           =   1575
         End
         Begin VB.OptionButton OptAceite 
            Caption         =   "Com Aceite"
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
            Index           =   2
            Left            =   5055
            TabIndex        =   7
            Top             =   285
            Width           =   1500
         End
      End
      Begin MSComCtl2.UpDown UpDownDataVencimento 
         Height          =   300
         Left            =   3135
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   352
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   300
         Left            =   1980
         TabIndex        =   13
         Top             =   352
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Totalizando 
         Height          =   300
         Left            =   1980
         TabIndex        =   14
         Top             =   811
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Títulos c/vcto até:"
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
         Left            =   285
         TabIndex        =   18
         Top             =   405
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Totalizando até (R$):"
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
         TabIndex        =   17
         Top             =   864
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         Left            =   4470
         TabIndex        =   16
         Top             =   405
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cliente:"
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
         Left            =   3795
         TabIndex        =   15
         Top             =   864
         Width           =   1365
      End
   End
   Begin VB.ComboBox Carteira 
      Height          =   315
      Left            =   5340
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   180
      Width           =   2400
   End
   Begin VB.ComboBox Cobrador 
      Height          =   315
      Left            =   2145
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   165
      Width           =   1920
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   2910
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3300
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1125
         Picture         =   "BorderoCobrancaX.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   330
         Left            =   90
         Picture         =   "BorderoCobrancaX.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   90
         Width           =   930
      End
   End
   Begin MSComCtl2.UpDown UpDownDataEmissao 
      Height          =   300
      Left            =   3300
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   705
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataEmissao 
      Height          =   300
      Left            =   2145
      TabIndex        =   20
      Top             =   705
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataContabil 
      Height          =   300
      Left            =   6525
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   705
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataContabil 
      Height          =   300
      Left            =   5355
      TabIndex        =   22
      Top             =   705
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Carteira:"
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
      Left            =   4560
      TabIndex        =   26
      Top             =   225
      Width           =   735
   End
   Begin VB.Label LabelDataContabil 
      AutoSize        =   -1  'True
      Caption         =   "Data Contábil:"
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
      Left            =   4080
      TabIndex        =   25
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Emissão:"
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
      Left            =   1335
      TabIndex        =   24
      Top             =   735
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cobrador:"
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
      Left            =   1260
      TabIndex        =   23
      Top             =   195
      Width           =   840
   End
End
Attribute VB_Name = "BorderoCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTBorderoCobranca
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTBorderoCobranca
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoSeguir_Click()
     Call objCT.BotaoSeguir_Click
End Sub

Private Sub Carteira_Click()
     Call objCT.Carteira_Click
End Sub

Private Sub Cobrador_Click()
     Call objCT.Cobrador_Click
End Sub

Function Trata_Parametros(Optional objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long
     Trata_Parametros = objCT.Trata_Parametros(objBorderoCobrancaEmissao)
End Function

Private Sub DataContabil_GotFocus()
     Call objCT.DataContabil_GotFocus
End Sub

Private Sub DataContabil_Validate(Cancel As Boolean)
     Call objCT.DataContabil_Validate(Cancel)
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
End Sub

Private Sub Estado_Validate(Cancel As Boolean)
     Call objCT.Estado_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub TipoCliente_Validate(Cancel As Boolean)
     Call objCT.TipoCliente_Validate(Cancel)
End Sub

Private Sub Totalizando_Validate(Cancel As Boolean)
     Call objCT.Totalizando_Validate(Cancel)
End Sub

Private Sub UpDownDataContabil_DownClick()
     Call objCT.UpDownDataContabil_DownClick
End Sub

Private Sub UpDownDataContabil_UpClick()
     Call objCT.UpDownDataContabil_UpClick
End Sub

Private Sub UpDownDataEmissao_DownClick()
     Call objCT.UpDownDataEmissao_DownClick
End Sub

Private Sub UpDownDataEmissao_UpClick()
     Call objCT.UpDownDataEmissao_UpClick
End Sub

Private Sub UpDownDataVencimento_DownClick()
     Call objCT.UpDownDataVencimento_DownClick
End Sub

Private Sub UpDownDataVencimento_UpClick()
     Call objCT.UpDownDataVencimento_UpClick
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub
Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub
Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub
Private Sub LabelDataContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataContabil, Source, X, Y)
End Sub
Private Sub LabelDataContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataContabil, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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



