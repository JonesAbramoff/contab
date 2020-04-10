VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ConfiguraTESOcx 
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   ScaleHeight     =   2655
   ScaleWidth      =   6195
   Begin VB.Frame Frame2 
      Caption         =   "Bloqueios"
      Height          =   1515
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   5940
      Begin VB.Frame Frame3 
         Caption         =   "Não permite incluir, alterar ou excluir movtos de conta corrente anteriores a"
         Height          =   660
         Left            =   120
         TabIndex        =   6
         Top             =   735
         Width           =   5685
         Begin MSMask.MaskEdBox DataBloqLimite 
            Height          =   315
            Left            =   1230
            TabIndex        =   7
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataBloqLimite 
            Height          =   300
            Left            =   2370
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Data Limite:"
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
            Left            =   180
            TabIndex        =   9
            Top             =   285
            Width           =   1065
         End
      End
      Begin VB.CheckBox BloqueioCTB 
         Caption         =   "Não permite incluir, alterar ou excluir movimentações de conta corrente se o período ou exercício contábil estiver fechado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   90
         TabIndex        =   5
         Top             =   165
         Width           =   5760
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   555
      Left            =   4905
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Width           =   1170
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ConfiguraTESOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "ConfiguraTESOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListaConfigura 
      Height          =   735
      ItemData        =   "ConfiguraTESOcx.ctx":02D8
      Left            =   120
      List            =   "ConfiguraTESOcx.ctx":02E5
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   225
      Width           =   4320
   End
End
Attribute VB_Name = "ConfiguraTESOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTConfiguraTES
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub ListaConfigura_ItemCheck(Item As Integer)
     Call objCT.ListaConfigura_ItemCheck(Item)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = objCT.Form_Load_Ocx()

End Function

Public Sub Form_UnLoad(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
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

Private Sub UserControl_Initialize()

    Set objCT = New CTConfiguraTES
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

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UpDownDataBloqLimite_DownClick()
     Call objCT.UpDownDataBloqLimite_DownClick
End Sub

Private Sub UpDownDataBloqLimite_UpClick()
     Call objCT.UpDownDataBloqLimite_UpClick
End Sub

Private Sub DataBloqLimite_Validate(Cancel As Boolean)
    Call objCT.DataBloqLimite_Validate(Cancel)
End Sub

Private Sub DataBloqLimite_Change()
    Call objCT.DataBloqLimite_Change
End Sub

Private Sub BloqueioCTB_Click()
    Call objCT.BloqueioCTB_Click
End Sub

