VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PrecoComponentes 
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   ScaleHeight     =   3705
   ScaleWidth      =   9435
   Begin VB.Frame Frame7 
      Caption         =   "Produtos Componentes"
      Height          =   2925
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   675
      Width           =   9240
      Begin MSMask.MaskEdBox PrecoComDesc 
         Height          =   315
         Left            =   8055
         TabIndex        =   4
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DescontoValor 
         Height          =   315
         Left            =   7050
         TabIndex        =   5
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PrecoSemDesc 
         Height          =   315
         Left            =   5235
         TabIndex        =   6
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DescontoPerc 
         Height          =   315
         Left            =   6255
         TabIndex        =   7
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PrecoUnitario 
         Height          =   315
         Left            =   4230
         TabIndex        =   8
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox UM 
         Height          =   315
         Left            =   3345
         TabIndex        =   9
         Top             =   210
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   285
         TabIndex        =   10
         Top             =   210
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   315
         Left            =   1815
         TabIndex        =   11
         Top             =   210
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridPrecos 
         Height          =   1920
         Left            =   -915
         TabIndex        =   12
         Top             =   390
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   3387
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
      End
      Begin VB.Label PrecoTotal 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   315
         Left            =   7500
         TabIndex        =   14
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Preço Total:"
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
         Left            =   6330
         TabIndex        =   13
         Top             =   2430
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7245
      ScaleHeight     =   495
      ScaleWidth      =   1965
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   2025
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1455
         Picture         =   "PrecoComponentes.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Caption         =   "Transportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   1245
      End
   End
End
Attribute VB_Name = "PrecoComponentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()

    lErro_Chama_Tela = SUCESSO

End Sub


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Preço Componentes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PrecoComponentes"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

