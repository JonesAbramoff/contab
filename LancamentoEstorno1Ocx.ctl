VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LancamentoEstorno1Ocx 
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   KeyPreview      =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   6030
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2415
      Picture         =   "LancamentoEstorno1Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Numeração Automática"
      Top             =   3345
      Width           =   300
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documento a ser Estornado"
      Height          =   1740
      Left            =   135
      TabIndex        =   17
      Top             =   120
      Width           =   5730
      Begin VB.Label Lote 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1245
         TabIndex        =   27
         Top             =   450
         Width           =   705
      End
      Begin VB.Label Origem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3975
         TabIndex        =   26
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
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
         Left            =   3225
         TabIndex        =   25
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   765
         TabIndex        =   24
         Top             =   465
         Width           =   450
      End
      Begin VB.Label Label9 
         Caption         =   "Exercício:"
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
         Left            =   360
         TabIndex        =   23
         Top             =   915
         Width           =   870
      End
      Begin VB.Label Exercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1245
         TabIndex        =   22
         Top             =   885
         Width           =   1530
      End
      Begin VB.Label Periodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3990
         TabIndex        =   21
         Top             =   885
         Width           =   1530
      End
      Begin VB.Label Label3 
         Caption         =   "Período:"
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
         Left            =   3120
         TabIndex        =   20
         Top             =   915
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   195
         TabIndex        =   19
         Top             =   1365
         Width           =   1020
      End
      Begin VB.Label Documento 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1260
         TabIndex        =   18
         Top             =   1320
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Documento de Estorno"
      Height          =   1845
      Left            =   120
      TabIndex        =   6
      Top             =   1965
      Width           =   5745
      Begin MSMask.MaskEdBox LoteEstorno 
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Top             =   375
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   2325
         TabIndex        =   7
         Top             =   870
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEstorno 
         Height          =   300
         Left            =   1170
         TabIndex        =   1
         Top             =   870
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DocumentoEstorno 
         Height          =   315
         Left            =   1185
         TabIndex        =   2
         Top             =   1365
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Caption         =   "Período:"
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
         Left            =   3150
         TabIndex        =   16
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   690
         TabIndex        =   15
         Top             =   900
         Width           =   480
      End
      Begin VB.Label PeriodoEstorno 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3975
         TabIndex        =   14
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label ExercicioEstorno 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3990
         TabIndex        =   13
         Top             =   900
         Width           =   1530
      End
      Begin VB.Label Label8 
         Caption         =   "Exercício:"
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
         Left            =   3030
         TabIndex        =   12
         Top             =   930
         Width           =   870
      End
      Begin VB.Label LabelLoteEstorno 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   750
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   435
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
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
         Left            =   3225
         TabIndex        =   10
         Top             =   480
         Width           =   660
      End
      Begin VB.Label OrigemEstorno 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3990
         TabIndex        =   9
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label LabelDocumentoEstorno 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   135
         TabIndex        =   8
         Top             =   1410
         Width           =   1020
      End
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      Height          =   555
      Left            =   3600
      Picture         =   "LancamentoEstorno1Ocx.ctx":00EA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1320
      Picture         =   "LancamentoEstorno1Ocx.ctx":01EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4110
      Width           =   975
   End
End
Attribute VB_Name = "LancamentoEstorno1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTLancamentoEstorno1
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoOK_Click()
     Call objCT.BotaoOK_Click
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub BotaoCancelar_Click()
     Call objCT.BotaoCancelar_Click
End Sub

Function Trata_Parametros(objLancamento_Cabecalho As ClassLancamento_Cabecalho, objBrowseConfigura As AdmBrowseConfigura) As Long
     Trata_Parametros = objCT.Trata_Parametros(objLancamento_Cabecalho, objBrowseConfigura)
End Function

Private Sub DataEstorno_GotFocus()
     Call objCT.DataEstorno_GotFocus
End Sub

Private Sub DataEstorno_Validate(Cancel As Boolean)
     Call objCT.DataEstorno_Validate(Cancel)
End Sub

Private Sub DocumentoEstorno_GotFocus()
     Call objCT.DocumentoEstorno_GotFocus
End Sub

Private Sub LoteEstorno_GotFocus()
     Call objCT.LoteEstorno_GotFocus
End Sub

Private Sub LoteEstorno_Validate(Cancel As Boolean)
     Call objCT.LoteEstorno_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub LabelLoteEstorno_Click()
     Call objCT.LabelLoteEstorno_Click
End Sub

Private Sub UpDown1_DownClick()
     Call objCT.UpDown1_DownClick
End Sub

Private Sub UpDown1_UpClick()
     Call objCT.UpDown1_UpClick
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

Private Sub UserControl_Initialize()
    Set objCT = New CTLancamentoEstorno1
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub



Private Sub Lote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Lote, Source, X, Y)
End Sub

Private Sub Lote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Lote, Button, Shift, X, Y)
End Sub

Private Sub Origem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Origem, Source, X, Y)
End Sub

Private Sub Origem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Origem, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Exercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Exercicio, Source, X, Y)
End Sub

Private Sub Exercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Exercicio, Button, Shift, X, Y)
End Sub

Private Sub Periodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Periodo, Source, X, Y)
End Sub

Private Sub Periodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Periodo, Button, Shift, X, Y)
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

Private Sub Documento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Documento, Source, X, Y)
End Sub

Private Sub Documento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Documento, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub PeriodoEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PeriodoEstorno, Source, X, Y)
End Sub

Private Sub PeriodoEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PeriodoEstorno, Button, Shift, X, Y)
End Sub

Private Sub ExercicioEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ExercicioEstorno, Source, X, Y)
End Sub

Private Sub ExercicioEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ExercicioEstorno, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub LabelLoteEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelLoteEstorno, Source, X, Y)
End Sub

Private Sub LabelLoteEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelLoteEstorno, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub OrigemEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OrigemEstorno, Source, X, Y)
End Sub

Private Sub OrigemEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OrigemEstorno, Button, Shift, X, Y)
End Sub

Private Sub LabelDocumentoEstorno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDocumentoEstorno, Source, X, Y)
End Sub

Private Sub LabelDocumentoEstorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDocumentoEstorno, Button, Shift, X, Y)
End Sub

