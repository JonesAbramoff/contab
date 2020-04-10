VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TransferenciaCaixaGeral 
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   ScaleHeight     =   5625
   ScaleWidth      =   9495
   Begin VB.ComboBox GerenteCod 
      Height          =   315
      Left            =   1095
      TabIndex        =   39
      Top             =   135
      Width           =   735
   End
   Begin VB.TextBox GerenteSenha 
      Height          =   315
      Left            =   1095
      TabIndex        =   38
      Top             =   540
      Width           =   2565
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8055
      ScaleHeight     =   495
      ScaleWidth      =   1140
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   90
      Width           =   1200
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   105
         Picture         =   "TransferenciaCaixaGeral.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   615
         Picture         =   "TransferenciaCaixaGeral.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Height          =   3750
      Index           =   3
      Left            =   150
      TabIndex        =   4
      Top             =   1650
      Width           =   9075
      Begin VB.ComboBox Ordem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5145
         TabIndex        =   33
         Top             =   60
         Width           =   1470
      End
      Begin VB.Frame FrameCartao 
         Caption         =   "Comprovantes de Venda"
         Height          =   3330
         Left            =   0
         TabIndex        =   6
         Top             =   390
         Width           =   9225
         Begin VB.ComboBox TipoParcelamento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7470
            TabIndex        =   12
            Top             =   315
            Width           =   765
         End
         Begin VB.ComboBox ECF 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8220
            TabIndex        =   11
            Top             =   300
            Width           =   840
         End
         Begin VB.CheckBox Selecionado 
            Height          =   195
            Left            =   30
            TabIndex        =   10
            Top             =   360
            Width           =   255
         End
         Begin VB.ComboBox Adm 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            TabIndex        =   9
            Top             =   330
            Width           =   1380
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Caption         =   "Marcar Todos"
            Height          =   585
            Left            =   225
            Picture         =   "TransferenciaCaixaGeral.ctx":06B0
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2640
            Width           =   1440
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Caption         =   "Desmarcar Todos"
            Height          =   585
            Left            =   1830
            Picture         =   "TransferenciaCaixaGeral.ctx":16CA
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2640
            Width           =   1440
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   4035
            TabIndex        =   13
            Top             =   345
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Doc 
            Height          =   300
            Left            =   2820
            TabIndex        =   14
            Top             =   345
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorCartao 
            Height          =   300
            Left            =   270
            TabIndex        =   15
            Top             =   345
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Autorizacao 
            Height          =   300
            Left            =   5205
            TabIndex        =   16
            Top             =   345
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumParcelas 
            Height          =   300
            Left            =   6390
            TabIndex        =   17
            Top             =   330
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCartoes 
            Height          =   1950
            Left            =   210
            TabIndex        =   18
            Top             =   585
            Width           =   8820
            _ExtentX        =   15558
            _ExtentY        =   3440
            _Version        =   393216
            Rows            =   7
            Cols            =   7
            FixedCols       =   0
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox CupomFiscal 
            Height          =   300
            Left            =   2730
            TabIndex        =   19
            Top             =   2715
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin VB.Label SaldoBoleto 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4935
            TabIndex        =   32
            Top             =   2745
            Width           =   1350
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
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
            Left            =   4335
            TabIndex        =   31
            Top             =   2775
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   6735
            TabIndex        =   30
            Top             =   2775
            Width           =   510
         End
         Begin VB.Label TotalSangriaBoleto 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7305
            TabIndex        =   29
            Top             =   2745
            Width           =   1350
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
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
            Left            =   705
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   120
            Width           =   450
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Nº Parcelas"
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
            Left            =   6420
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   135
            Width           =   1020
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Doc"
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
            Left            =   3120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   150
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
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
            Left            =   4185
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   135
            Width           =   705
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Autorização"
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
            Left            =   5295
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   24
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cupom Fiscal"
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
            Left            =   2775
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   23
            Top             =   2520
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ECF"
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
            Left            =   8445
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   22
            Top             =   105
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Parcto"
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
            Left            =   7545
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   21
            Top             =   135
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Administradora"
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
            Left            =   1530
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            Top             =   150
            Width           =   1260
         End
      End
      Begin VB.CommandButton BotaoTransferenciaBoletos 
         Caption         =   "Transfere"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7155
         TabIndex        =   5
         Top             =   45
         Width           =   1830
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Ordenação:"
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
         Left            =   4080
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   105
         Width           =   1005
      End
   End
   Begin VB.Frame FrameOp 
      BorderStyle     =   0  'None
      Height          =   3750
      Index           =   1
      Left            =   150
      TabIndex        =   0
      Top             =   1665
      Width           =   9075
      Begin VB.CommandButton BotaoTransfereDinheiro 
         Caption         =   "Transfere"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5280
         TabIndex        =   1
         Top             =   2610
         Width           =   3345
      End
      Begin MSMask.MaskEdBox SangriaDinheiro 
         Height          =   645
         Left            =   2865
         TabIndex        =   2
         Top             =   960
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1138
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "A9"
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1305
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   3
         Top             =   990
         Width           =   1380
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4485
      Left            =   60
      TabIndex        =   40
      Top             =   1065
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   7911
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dinheiro"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cheques"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Boletos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vales/Tickets"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
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
      Left            =   390
      TabIndex        =   43
      Top             =   555
      Width           =   615
   End
   Begin VB.Label GerenteNomeRed 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1980
      TabIndex        =   42
      Top             =   150
      Width           =   1665
   End
   Begin VB.Label Label1 
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
      Left            =   270
      TabIndex        =   41
      Top             =   195
      Width           =   750
   End
End
Attribute VB_Name = "TransferenciaCaixaGeral"
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
    Caption = "Transferência Caixa Geral"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TransferenciaCaixaGeral"
    
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


