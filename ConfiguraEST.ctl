VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ConfiguraEST 
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   ScaleHeight     =   6135
   ScaleWidth      =   6135
   Begin VB.Frame Frame2 
      Caption         =   "Bloqueios"
      Height          =   1515
      Left            =   90
      TabIndex        =   19
      Top             =   4590
      Width           =   5940
      Begin VB.CheckBox BloqueioCTB 
         Caption         =   "Não permite incluir, alterar ou excluir movimentações de estoque se o período ou exercício contábil estiver fechado"
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
         TabIndex        =   24
         Top             =   165
         Width           =   5760
      End
      Begin VB.Frame Frame3 
         Caption         =   "Não permite incluir, alterar ou excluir movtos de estoque anteriores a"
         Height          =   660
         Left            =   120
         TabIndex        =   20
         Top             =   735
         Width           =   5685
         Begin MSMask.MaskEdBox DataBloqLimite 
            Height          =   315
            Left            =   1230
            TabIndex        =   21
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
            TabIndex        =   22
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
            TabIndex        =   23
            Top             =   285
            Width           =   1065
         End
      End
   End
   Begin VB.CheckBox GeraReqCompraEmLote 
      Caption         =   "Gerar as Requisições de Compra por Lote"
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
      Left            =   225
      TabIndex        =   5
      Top             =   2025
      Width           =   4110
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   90
      TabIndex        =   12
      Top             =   2895
      Width           =   5940
      Begin VB.CommandButton DownPrioridade 
         Height          =   660
         Left            =   4665
         Picture         =   "ConfiguraEST.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   870
         Width           =   285
      End
      Begin VB.CommandButton UpPrioridade 
         Height          =   660
         Left            =   4665
         Picture         =   "ConfiguraEST.ctx":01F2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   285
      End
      Begin VB.TextBox Prioridade 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1350
         TabIndex        =   13
         Top             =   885
         Width           =   3225
      End
      Begin MSFlexGridLib.MSFlexGrid GridPrioridades 
         Height          =   1380
         Left            =   885
         TabIndex        =   7
         Top             =   180
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   2434
         _Version        =   393216
      End
   End
   Begin VB.CheckBox AceitaQtdNegativa 
      Caption         =   "Aceita quantidades negativas no estoque"
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
      Left            =   225
      TabIndex        =   4
      Top             =   1680
      Width           =   4020
   End
   Begin VB.CheckBox IncluiFrete 
      Caption         =   "Inclui Frete e outras despesas no cálculo do custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   225
      TabIndex        =   3
      Top             =   1410
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4725
      ScaleHeight     =   495
      ScaleWidth      =   1185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   165
      Width           =   1245
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   645
         Picture         =   "ConfiguraEST.ctx":03E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ConfiguraEST.ctx":0562
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.ListBox ListaConfigura 
      Height          =   735
      ItemData        =   "ConfiguraEST.ctx":06BC
      Left            =   180
      List            =   "ConfiguraEST.ctx":06C9
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   165
      Width           =   4320
   End
   Begin MSMask.MaskEdBox IntervaloProducao 
      Height          =   315
      Left            =   5310
      TabIndex        =   2
      Top             =   1065
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ClasseUMTempo 
      Height          =   315
      Left            =   4455
      TabIndex        =   6
      Top             =   2610
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataInicio 
      Height          =   315
      Left            =   2835
      TabIndex        =   15
      Top             =   2310
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataInicio 
      Height          =   300
      Left            =   3975
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2310
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Data Início Operações MRP:"
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
      Left            =   195
      TabIndex        =   14
      Top             =   2385
      Width           =   2625
   End
   Begin VB.Label LabelClasseUMTempo 
      Caption         =   "Classe de Unidade de Medida de Tempo Padrão:"
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
      Left            =   195
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   2685
      Width           =   4200
   End
   Begin VB.Label LabeIntervaloProducao 
      Caption         =   "Intervalo médio em dias entre a produção dos insumos e a produção da mercadoria que utiliza os insumos produzidos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   180
      TabIndex        =   10
      Top             =   990
      Width           =   5040
   End
End
Attribute VB_Name = "ConfiguraEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTConfiguraEST
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

Private Sub DataInicio_Validate(Cancel As Boolean)
    Call objCT.DataInicio_Validate(Cancel)
End Sub

Private Sub IntervaloProducao_Change()
    Call objCT.IntervaloProducao_Change
End Sub

Private Sub IncluiFrete_Click()
    Call objCT.IncluiFrete_Click
End Sub
Private Sub AceitaQtdNegativa_Click()
    Call objCT.AceitaQtdNegativa_Click
End Sub

Private Sub LabelClasseUMTempo_Click()
    Call objCT.LabelClasseUMTempo_Click
End Sub

Private Sub ListaConfigura_ItemCheck(Item As Integer)
     Call objCT.ListaConfigura_ItemCheck(Item)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
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
    Set objCT = New CTConfiguraEST
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

'Incluido por Jorge Specian - 06/04/2005
'---------------------------------------
Private Sub GridPrioridades_Click()
     Call objCT.GridPrioridades_Click
End Sub

Private Sub GridPrioridades_EnterCell()
     Call objCT.GridPrioridades_EnterCell
End Sub

Private Sub GridPrioridades_GotFocus()
     Call objCT.GridPrioridades_GotFocus
End Sub

Private Sub GridPrioridades_KeyPress(KeyAscii As Integer)
     Call objCT.GridPrioridades_KeyPress(KeyAscii)
End Sub

Private Sub GridPrioridades_LeaveCell()
     Call objCT.GridPrioridades_LeaveCell
End Sub

Private Sub GridPrioridades_Validate(Cancel As Boolean)
     Call objCT.GridPrioridades_Validate(Cancel)
End Sub

Private Sub GridPrioridades_RowColChange()
     Call objCT.GridPrioridades_RowColChange
End Sub

Private Sub GridPrioridades_Scroll()
     Call objCT.GridPrioridades_Scroll
End Sub

Private Sub Prioridade_Change()
     Call objCT.Prioridade_Change
End Sub

Private Sub Prioridade_GotFocus()
     Call objCT.Prioridade_GotFocus
End Sub

Private Sub Prioridade_KeyPress(KeyAscii As Integer)
     Call objCT.Prioridade_KeyPress(KeyAscii)
End Sub

Private Sub Prioridade_Validate(Cancel As Boolean)
     Call objCT.Prioridade_Validate(Cancel)
End Sub

Private Sub DownPrioridade_Click()
     Call objCT.DownPrioridade_Click
End Sub

Private Sub UpPrioridade_Click()
     Call objCT.UpPrioridade_Click
End Sub

Private Sub UpDownDataInicio_DownClick()
     Call objCT.UpDownDataInicio_DownClick
End Sub

Private Sub UpDownDataInicio_UpClick()
     Call objCT.UpDownDataInicio_UpClick
End Sub

Private Sub GeraReqCompraEmLote_Click()
    Call objCT.GeraReqCompraEmLote_Click
End Sub
'-------------------------------------------

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
