VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ConfiguraCROcx 
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   ScaleHeight     =   5625
   ScaleWidth      =   7380
   Begin VB.Frame Frame6 
      Caption         =   "Localização de Arquivos"
      Height          =   675
      Left            =   120
      TabIndex        =   27
      Top             =   4890
      Width           =   7170
      Begin VB.TextBox NomeDiretorioBoleto 
         Height          =   315
         Left            =   1065
         TabIndex        =   10
         Top             =   225
         Width           =   5430
      End
      Begin VB.CommandButton BotaoProcurarBoleto 
         Caption         =   "..."
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
         Left            =   6495
         TabIndex        =   11
         Top             =   195
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Boleto:"
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
         Left            =   330
         TabIndex        =   28
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Fatura"
      Height          =   945
      Left            =   4560
      TabIndex        =   24
      Top             =   780
      Width           =   2715
      Begin MSMask.MaskEdBox NumFatura 
         Height          =   300
         Left            =   1590
         TabIndex        =   25
         Top             =   390
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99999999"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Próximo Número:"
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
         Left            =   150
         TabIndex        =   26
         Top             =   450
         Width           =   1440
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cobrança"
      Height          =   945
      Left            =   135
      TabIndex        =   23
      Top             =   780
      Width           =   4365
      Begin VB.ComboBox ComboFilialCobr 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   510
         Width           =   2385
      End
      Begin VB.OptionButton OptionCobranca 
         Caption         =   "Independente Por Filial"
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
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Width           =   2520
      End
      Begin VB.OptionButton OptionCobranca 
         Caption         =   "Centralizada em:"
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
         Left            =   135
         TabIndex        =   2
         Top             =   570
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   555
      Left            =   6105
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   180
      Width           =   1170
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ConfiguraCROcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "ConfiguraCROcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListaConfigura 
      Height          =   510
      ItemData        =   "ConfiguraCROcx.ctx":02D8
      Left            =   255
      List            =   "ConfiguraCROcx.ctx":02E2
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   180
      Width           =   4320
   End
   Begin VB.Frame Frame1 
      Caption         =   "Padrões do Sistema"
      Height          =   3045
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   7170
      Begin VB.Frame Frame2 
         Caption         =   "Para Atrasos de Pagamento"
         Height          =   900
         Left            =   270
         TabIndex        =   16
         Top             =   195
         Width           =   6660
         Begin VB.Frame Frame3 
            Caption         =   "Juros"
            Height          =   645
            Left            =   2130
            TabIndex        =   19
            Top             =   165
            Width           =   4365
            Begin MSMask.MaskEdBox JurosMensais 
               Height          =   300
               Left            =   1080
               TabIndex        =   5
               Top             =   240
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "#0.#0\%"
               PromptChar      =   "_"
            End
            Begin VB.Label Label10 
               Caption         =   "Diários:"
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
               Left            =   2190
               TabIndex        =   22
               Top             =   270
               Width           =   765
            End
            Begin VB.Label Label4 
               Caption         =   "Mensais:"
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
               Left            =   240
               TabIndex        =   21
               Top             =   270
               Width           =   795
            End
            Begin VB.Label JurosDiarios 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2970
               TabIndex        =   20
               Top             =   240
               Width           =   900
            End
         End
         Begin MSMask.MaskEdBox PercMulta 
            Height          =   285
            Left            =   765
            TabIndex        =   4
            Top             =   375
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "#0.#0\%"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Multa:"
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
            Left            =   150
            TabIndex        =   17
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Descontos por Antecipação de Pagamento"
         Height          =   1890
         Left            =   270
         TabIndex        =   18
         Top             =   1110
         Width           =   6660
         Begin VB.ComboBox TipoDesconto 
            Height          =   315
            ItemData        =   "ConfiguraCROcx.ctx":0346
            Left            =   855
            List            =   "ConfiguraCROcx.ctx":0348
            TabIndex        =   6
            Top             =   435
            Width           =   1890
         End
         Begin MSMask.MaskEdBox Dias 
            Height          =   225
            Left            =   2745
            TabIndex        =   7
            Top             =   480
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualDesc 
            Height          =   225
            Left            =   3480
            TabIndex        =   8
            Top             =   450
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDescontos 
            Height          =   1110
            Left            =   765
            TabIndex        =   9
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1958
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
End
Attribute VB_Name = "ConfiguraCROcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTConfiguraCR
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

Private Sub Dias_Change()
     Call objCT.Dias_Change
End Sub

Private Sub Dias_GotFocus()
     Call objCT.Dias_GotFocus
End Sub

Private Sub Dias_KeyPress(KeyAscii As Integer)
     Call objCT.Dias_KeyPress(KeyAscii)
End Sub

Private Sub Dias_Validate(Cancel As Boolean)
     Call objCT.Dias_Validate(Cancel)
End Sub

Private Sub ListaConfigura_ItemCheck(Item As Integer)
     Call objCT.ListaConfigura_ItemCheck(Item)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub NumFatura_Change()
    Call objCT.NumFatura_Change
End Sub

Private Sub NumFatura_GotFocus()
    Call objCT.NumFatura_GotFocus
End Sub

Private Sub OptionCobranca_Click(Index As Integer)
     Call objCT.OptionCobranca_Click(Index)
End Sub

Private Sub PercentualDesc_Change()
     Call objCT.PercentualDesc_Change
End Sub

Private Sub PercentualDesc_GotFocus()
     Call objCT.PercentualDesc_GotFocus
End Sub

Private Sub PercentualDesc_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualDesc_KeyPress(KeyAscii)
End Sub

Private Sub PercentualDesc_Validate(Cancel As Boolean)
     Call objCT.PercentualDesc_Validate(Cancel)
End Sub

Private Sub JurosMensais_Change()
     Call objCT.JurosMensais_Change
End Sub

Private Sub JurosMensais_Validate(Cancel As Boolean)
     Call objCT.JurosMensais_Validate(Cancel)
End Sub

Private Sub PercMulta_Change()
     Call objCT.PercMulta_Change
End Sub

Private Sub PercMulta_Validate(Cancel As Boolean)
     Call objCT.PercMulta_Validate(Cancel)
End Sub

Private Sub TipoDesconto_Change()
     Call objCT.TipoDesconto_Change
End Sub

Private Sub TipoDesconto_GotFocus()
     Call objCT.TipoDesconto_GotFocus
End Sub

Private Sub TipoDesconto_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto_Validate(Cancel)
End Sub

Private Sub GridDescontos_Click()
     Call objCT.GridDescontos_Click
End Sub

Private Sub GridDescontos_EnterCell()
     Call objCT.GridDescontos_EnterCell
End Sub

Private Sub GridDescontos_GotFocus()
     Call objCT.GridDescontos_GotFocus
End Sub

Private Sub GridDescontos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridDescontos_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridDescontos_KeyPress(KeyAscii As Integer)
     Call objCT.GridDescontos_KeyPress(KeyAscii)
End Sub

Private Sub GridDescontos_LeaveCell()
     Call objCT.GridDescontos_LeaveCell
End Sub

Private Sub GridDescontos_Validate(Cancel As Boolean)
     Call objCT.GridDescontos_Validate(Cancel)
End Sub

Private Sub GridDescontos_RowColChange()
     Call objCT.GridDescontos_RowColChange
End Sub

Private Sub GridDescontos_Scroll()
     Call objCT.GridDescontos_Scroll
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

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
    Set objCT = New CTConfiguraCR
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


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub JurosDiarios_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(JurosDiarios, Source, X, Y)
End Sub

Private Sub JurosDiarios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(JurosDiarios, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub NomeDiretorioBoleto_Validate(Cancel As Boolean)
    Call objCT.NomeDiretorioBoleto_Validate(Cancel)
End Sub

Private Sub NomeDiretorioBoleto_Change()
    Call objCT.NomeDiretorioBoleto_Change
End Sub

Private Sub BotaoProcurarBoleto_Click()
    Call objCT.BotaoProcurarBoleto_Click
End Sub
