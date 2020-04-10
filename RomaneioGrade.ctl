VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl RomaneioGrade 
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   KeyPreview      =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   8340
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   4170
      Picture         =   "RomaneioGrade.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7890
      Width           =   1005
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2760
      Picture         =   "RomaneioGrade.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7890
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Produto "
      Height          =   870
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   8145
      Begin VB.Label UnidadeMed 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5100
         TabIndex        =   30
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UM:"
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
         Index           =   8
         Left            =   4710
         TabIndex        =   29
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Top             =   555
         Width           =   930
      End
      Begin VB.Label DescricaoPai 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1590
         TabIndex        =   15
         Top             =   495
         Width           =   6345
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   6
         Left            =   900
         TabIndex        =   14
         Top             =   210
         Width           =   660
      End
      Begin VB.Label CodigoPai 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1590
         TabIndex        =   13
         Top             =   150
         Width           =   2685
      End
   End
   Begin VB.Frame FrameGrade 
      Caption         =   "Grade"
      Height          =   4725
      Left            =   120
      TabIndex        =   17
      Top             =   945
      Width           =   8145
      Begin VB.Frame FrameQuantidades 
         Caption         =   "Quantidades"
         Height          =   780
         Left            =   165
         TabIndex        =   22
         Top             =   4275
         Width           =   7860
         Begin MSMask.MaskEdBox QuantCancelada 
            Height          =   300
            Left            =   720
            TabIndex        =   5
            Top             =   390
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Label QUantCancelLbl 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   720
            TabIndex        =   28
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label LabelCabcelda 
            AutoSize        =   -1  'True
            Caption         =   "Cancelada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   1005
            TabIndex        =   27
            Top             =   180
            Width           =   915
         End
         Begin VB.Label QuantFaturada 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3195
            TabIndex        =   26
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label LabelFaturada 
            AutoSize        =   -1  'True
            Caption         =   "Faturada"
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
            Left            =   3540
            TabIndex        =   25
            Top             =   165
            Width           =   765
         End
         Begin VB.Label QuantReservada 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5700
            TabIndex        =   24
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label LabelReservada 
            AutoSize        =   -1  'True
            Caption         =   "Reservada"
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
            Left            =   5940
            TabIndex        =   23
            Top             =   165
            Width           =   930
         End
      End
      Begin VB.Frame FrameProd 
         Height          =   855
         Left            =   165
         TabIndex        =   43
         Top             =   4275
         Visible         =   0   'False
         Width           =   7860
         Begin VB.TextBox CodOPProd 
            Height          =   300
            Left            =   1905
            MaxLength       =   6
            TabIndex        =   44
            Top             =   150
            Width           =   1035
         End
         Begin MSMask.MaskEdBox AlmoxProd 
            Height          =   300
            Left            =   5400
            TabIndex        =   45
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteProd 
            Height          =   300
            Left            =   5400
            TabIndex        =   51
            Top             =   480
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HorasMaqProd 
            Height          =   300
            Left            =   1905
            TabIndex        =   49
            Top             =   480
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            _Version        =   393216
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
         Begin VB.Label LabelHorasMaqProd 
            AutoSize        =   -1  'True
            Caption         =   "Horas Máquina:"
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
            Left            =   480
            TabIndex        =   50
            Top             =   510
            Width           =   1350
         End
         Begin VB.Label LabelAmoxProd 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado:"
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
            TabIndex        =   48
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label LabelOPProd 
            AutoSize        =   -1  'True
            Caption         =   "Ordem Produção:"
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
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   47
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label LabelLoteProd 
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
            Height          =   195
            Left            =   4890
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   46
            Top             =   525
            Width           =   450
         End
      End
      Begin VB.Frame FrameOP 
         Height          =   570
         Left            =   165
         TabIndex        =   32
         Top             =   4290
         Visible         =   0   'False
         Width           =   7860
         Begin VB.ComboBox Versao 
            Height          =   315
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   165
            Width           =   1815
         End
         Begin MSMask.MaskEdBox AlmoxOP 
            Height          =   315
            Left            =   4680
            TabIndex        =   3
            Top             =   165
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelAlmoxOP 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3420
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   34
            Top             =   210
            Width           =   1155
         End
         Begin VB.Label LabelVersao 
            AutoSize        =   -1  'True
            Caption         =   "Versão:"
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
            Height          =   225
            Left            =   525
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   33
            Top             =   210
            Width           =   660
         End
      End
      Begin VB.Frame FrameAlmoxarifado 
         Height          =   525
         Left            =   135
         TabIndex        =   6
         Top             =   4305
         Visible         =   0   'False
         Width           =   7875
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   300
            Left            =   2730
            TabIndex        =   4
            Top             =   150
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelAlmoxarifado 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado:"
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
            Left            =   1485
            TabIndex        =   31
            Top             =   225
            Width           =   1155
         End
      End
      Begin VB.Frame FrameProdSai 
         Height          =   1275
         Left            =   135
         TabIndex        =   35
         Top             =   4260
         Visible         =   0   'False
         Width           =   7860
         Begin VB.ComboBox FilialOPProdSai 
            Height          =   315
            Left            =   4320
            TabIndex        =   11
            Top             =   525
            Width           =   2160
         End
         Begin VB.TextBox CodOPProdSai 
            Height          =   300
            Left            =   4350
            MaxLength       =   6
            TabIndex        =   8
            Top             =   180
            Width           =   960
         End
         Begin MSMask.MaskEdBox AlmoxProdSai 
            Height          =   300
            Left            =   1320
            TabIndex        =   7
            Top             =   165
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteProdSai 
            Height          =   300
            Left            =   1320
            TabIndex        =   10
            Top             =   510
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdOPProdSai 
            Height          =   300
            Left            =   6630
            TabIndex        =   9
            Top             =   165
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label QuantDispProdSai 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4335
            TabIndex        =   42
            Top             =   885
            Width           =   1425
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade Disponível:"
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
            Left            =   2205
            TabIndex        =   41
            Top             =   930
            Width           =   2025
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Filial O.P.:"
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
            Left            =   3390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   570
            Width           =   900
         End
         Begin VB.Label LabelLoteProdSai 
            AutoSize        =   -1  'True
            Caption         =   "Lote/O.P.:"
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
            Left            =   375
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   555
            Width           =   915
         End
         Begin VB.Label LabelProdOPProdSai 
            AutoSize        =   -1  'True
            Caption         =   "Produto O.P.:"
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
            Left            =   5430
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label LabelOPProdSai 
            AutoSize        =   -1  'True
            Caption         =   "Ordem Produção:"
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
            Left            =   2805
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   37
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label LabelAlmoxProdSai 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado:"
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   36
            Top             =   210
            Width           =   1155
         End
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   270
         Index           =   0
         Left            =   3060
         TabIndex        =   0
         Top             =   1200
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridGrade 
         Height          =   2850
         Left            =   120
         TabIndex        =   1
         Top             =   195
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   5027
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComDlg.CommonDialog CommonDialog 
         Left            =   6270
         Top             =   3450
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Escolhendo Figura para o Produto"
      End
      Begin VB.Image Figura 
         BorderStyle     =   1  'Fixed Single
         Height          =   1200
         Left            =   6765
         Stretch         =   -1  'True
         Top             =   3420
         Width           =   1200
      End
      Begin VB.Label LabelDescricaoGrade 
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
         Height          =   195
         Left            =   870
         TabIndex        =   21
         Top             =   3900
         Width           =   915
      End
      Begin VB.Label DescricaoFilho 
         BorderStyle     =   1  'Fixed Single
         Height          =   555
         Left            =   1875
         TabIndex        =   20
         Top             =   3855
         Width           =   4770
      End
      Begin VB.Label LabelProdutoGrade 
         AutoSize        =   -1  'True
         Caption         =   "Produto Grade:"
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
         Left            =   525
         TabIndex        =   19
         Top             =   3480
         Width           =   1305
      End
      Begin VB.Label CodigoFilho 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1875
         TabIndex        =   18
         Top             =   3465
         Width           =   2535
      End
   End
End
Attribute VB_Name = "RomaneioGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoLocalizacao As AdmEvento
Attribute objEventoLocalizacao.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoEstoqueProdSai As AdmEvento
Attribute objEventoEstoqueProdSai.VB_VarHelpID = -1
Private WithEvents objEventoOP As AdmEvento
Attribute objEventoOP.VB_VarHelpID = -1
Private WithEvents objEventoProdutoOP As AdmEvento
Attribute objEventoProdutoOP.VB_VarHelpID = -1
Private WithEvents objEventoOPProd As AdmEvento
Attribute objEventoOPProd.VB_VarHelpID = -1
Private WithEvents objEventoEstoqueProd As AdmEvento
Attribute objEventoEstoqueProd.VB_VarHelpID = -1

Dim gobjRomaneioGrade As ClassRomaneioGrade
Dim dQuantCanceladaAnterior As Double

Dim gobjGradeLinha As ClassGradeLinCol
Dim gobjGradeColuna As ClassGradeLinCol

Dim iTeclaAnterior As Integer

'Grid Categoria
Dim objGridGrade As AdmGrid

Public iAlterado As Integer

'**** inicio do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Romaneio - Grade"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RomaneioGrade"

End Function

Public Sub Show()
'    Me.Show
'    Parent.SetFocus
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

Private Sub Almoxarifado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Almoxarifado_Validate

    If Len(Trim(CodigoFilho.Caption)) = 0 Then Exit Sub

    'Se o Almoxarifado está preenchido
    If Len(Trim(Almoxarifado.Text)) > 0 Then

        'Valida o ALmoxarifado
        lErro = TP_Almoxarifado_Filial_Produto_Grid(gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 39186
        'Se não for encontrado --> Erro
        If lErro = 25157 Then gError 39187
        If lErro = 25162 Then gError 39188

        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = objAlmoxarifado.iCodigo
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = objAlmoxarifado.sNomeReduzido

    Else
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = 0
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = ""
    
    End If

    Exit Sub

Erro_Almoxarifado_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 39187
            'Pergunta de deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)
            'Se a resposta for sim
            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                'Chama a Tela Almoxarifados
                Call Chama_Tela_Modal("Almoxarifado", objAlmoxarifado)


            End If

        Case 39188

            'Pergunta se deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", Codigo_Extrai(Almoxarifado.Text))

            'Se a resposta for positiva
            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = Codigo_Extrai(Almoxarifado.Text)

                'Chama a tela de Almoxarifados
                Call Chama_Tela_Modal("Almoxarifado", objAlmoxarifado)

            End If
            
        Case 39186

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174175)

    End Select

    Exit Sub

End Sub

Private Sub CodOPProdSai_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodOPProdSai_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemOP As New ClassItemOP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iCont As Integer
Dim objItemOPUnico As ClassItemOP
Dim sProdutoOPEnxuto As String

On Error GoTo Erro_CodOpProdSai_Validate

    If Len(Trim(CodOPProdSai.Text)) > 0 Then

        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = CodOPProdSai.Text

        lErro = CF("OrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 126595

        If lErro = 30368 Then gError 126596

        'ordem de producao baixada
        If lErro = 55316 Then gError 126597

        lErro = CF("Produto_Formata", ProdOPProdSai.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 126598

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            objItemOP.iFilialEmpresa = giFilialEmpresa
            objItemOP.sCodigo = CodOPProdSai.Text
            objItemOP.sProduto = sProdutoFormatado

            lErro = CF("ItemOP_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError 126599

            If lErro = 34711 Then gError 126600
            
        Else
        
            lErro = CF("ItensOrdemProducao_Le", objOrdemProducao)
            If lErro <> SUCESSO And lErro <> 30401 Then gError 126601
            
            For Each objItemOP In objOrdemProducao.colItens
                If objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
                    If objItemOPUnico Is Nothing Then Set objItemOPUnico = objItemOP
                    iCont = iCont + 1
                End If
            Next
            
            If iCont = 1 Then
                'Mascara o ProdutoOP
                lErro = Mascara_RetornaProdutoEnxuto(objItemOPUnico.sProduto, sProdutoOPEnxuto)
                If lErro <> SUCESSO Then gError 126602
        
                ProdOPProdSai.PromptInclude = False
                ProdOPProdSai.Text = sProdutoOPEnxuto
                ProdOPProdSai.PromptInclude = True
        
                gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProdOP = objItemOPUnico.sProduto
        
            End If
        End If
          
    Else
        
        ProdOPProdSai.PromptInclude = False
        ProdOPProdSai.Text = ""
        ProdOPProdSai.PromptInclude = True

        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProdOP = ""

    End If

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sCodOP = CodOPProdSai.Text
          
    Exit Sub

Erro_CodOpProdSai_Validate:

    Cancel = True

    Select Case gErr

        Case 126595, 126598, 126599, 126601, 126602

        Case 126596
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OPCODIGO_NAO_CADASTRADO", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then

                Call Chama_Tela_Modal("OrdemProducao", objOrdemProducao)

            End If

        Case 126597
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
        
        Case 126600
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174176)

    End Select

    Exit Sub

End Sub

Private Sub LabelOPProdSai_Click()

Dim objOrdemProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection

On Error GoTo Erro_LabelOPProdSai_Click
    
    Call Chama_Tela_Modal("OrdemProducaoLista", colSelecao, objOrdemProducao, objEventoOP)
   
    Exit Sub

Erro_LabelOPProdSai_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174177)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdOPProdSai_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objItemOP As New ClassItemOP

On Error GoTo Erro_LabelProdOPProdSai_Click

    ' Formata o ProdutoOp
    lErro = CF("Produto_Formata", ProdOPProdSai.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 126607
    
    'se a OP estiver preenchida, mostra só os produtos da OP em questão
    If Len(Trim(CodOPProdSai.Text)) = 0 Then gError 126608
        
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objItemOP.sProduto = sProdutoFormatado
    
    colSelecao.Add Trim(CodOPProdSai.Text)
    
    Call Chama_Tela_Modal("ItemOrdemProducao_OPLista", colSelecao, objItemOP, objEventoProdutoOP)
        
    Exit Sub
    
Erro_LabelProdOPProdSai_Click:

    Select Case gErr
        
        Case 126607
        
        Case 126608
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO_OP", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174178)
        
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProdutoOP_evSelecao(obj1 As Object)

Dim objItemOP As ClassItemOP
Dim lErro As Long
Dim sProdutoMascarado As String
Dim objItemR As ClassItemRomaneioGrade

On Error GoTo Erro_objEventoProdutoOP_evSelecao

    Set objItemOP = obj1

    lErro = Mascara_MascararProduto(objItemOP.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 126609

    ProdOPProdSai.PromptInclude = False
    ProdOPProdSai.Text = sProdutoMascarado
    ProdOPProdSai.PromptInclude = True

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProdOP = objItemOP.sProduto
            
    Me.Show

    Exit Sub

Erro_objEventoProdutoOP_evSelecao:

    Select Case gErr

        Case 126609
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemOP.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174179)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoOP_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOrdemDeProducao
    
On Error GoTo Erro_objEventoOP_evSelecao

    Set objOrdemProducao = obj1
    
    CodOPProdSai.Text = objOrdemProducao.sCodigo
        
    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sCodOP = objOrdemProducao.sCodigo
        
    Me.Show

    Exit Sub
    
Erro_objEventoOP_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174180)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub ProdOPProdSai_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ProdOPProdSai_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objItemOP As New ClassItemOP
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_ProdOPProdSai_Validate

    If Len(Trim(ProdOPProdSai.ClipText)) > 0 Then

        lErro = CF("Produto_Critica", ProdOPProdSai.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 126603

        If lErro = 25041 Then gError 126604

        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProdOP = objProduto.sCodigo

        If Len(Trim(CodOPProdSai.Text)) > 0 Then

            objItemOP.sCodigo = CodOPProdSai.Text
            objItemOP.iFilialEmpresa = giFilialEmpresa

            objItemOP.sProduto = objProduto.sCodigo

            lErro = CF("ItemOP_Le", objItemOP)
            If lErro <> SUCESSO And lErro <> 34711 Then gError 126605
            
            If lErro = 34711 Then gError 126606

        End If

    Else
    
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProdOP = ""

    End If

    Exit Sub

Erro_ProdOPProdSai_Validate:

    Cancel = True

    Select Case gErr

        Case 126603, 126605
        
        Case 126604
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdOPProdSai.Text)
            If vbMsg = vbYes Then
                objProduto.sCodigo = ProdOPProdSai.Text

                Call Chama_Tela_Modal("Produto", objProduto)
            End If

        Case 126606
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174181)

    End Select

    Exit Sub

End Sub

Private Sub LoteProdSai_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LoteProdSai_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim iEntradaSaida As Integer

On Error GoTo Erro_LoteProdSai_Validate

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sLote = LoteProdSai.Text

    If Len(Trim(LoteProdSai.Text)) > 0 Then
        
        objProduto.sCodigo = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 126624
            
        If lErro = 28030 Then gError 126625
                
        'Se for rastro por lote
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then

            objRastroLote.sCodigo = LoteProdSai.Text
            objRastroLote.sProduto = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto

            'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
            lErro = CF("RastreamentoLote_Le", objRastroLote)
            If lErro <> SUCESSO And lErro <> 75710 Then gError 126626

            'Se não encontrou --> Erro
            If lErro = 75710 Then gError 126627

            'Preenche a Quantidade do Lote
            lErro = QuantLote_Calcula(QuantDispProdSai)
            If lErro <> SUCESSO Then gError 126636
                
        'Se for rastro por OP
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then

            If Len(Trim(FilialOPProdSai.Text)) > 0 Then
                
                objRastroLote.sCodigo = LoteProdSai.Text
                objRastroLote.sProduto = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
                objRastroLote.iFilialOP = StrParaInt(FilialOPProdSai.Text)

                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 126628

                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 126629

                'Preenche a Quantidade do Lote
                lErro = QuantLote_Calcula(QuantDispProdSai)
                If lErro <> SUCESSO Then gError 126641

            End If
                            
        Else

            'Preenche a Quantidade do Lote
            lErro = QuantDisponivel_Calcula1(QuantDispProdSai)
            If lErro <> SUCESSO Then gError 126642
        
        End If
                            
    End If
            
    If Len(Trim(QuantDispProdSai.Caption)) <> 0 Then

        lErro = Testa_QuantRequisitada(QuantDispProdSai)
        If lErro <> SUCESSO Then gError 126630

    End If

    Exit Sub

Erro_LoteProdSai_Validate:

    Cancel = True

    Select Case gErr

        Case 126624, 126626, 126628, 126630, 126631, 126636
        
        Case 126625
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 126627, 126629
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then Call Chama_Tela_Modal("RastreamentoLote", objRastroLote)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174182)

    End Select

    Exit Sub

End Sub

Private Sub FilialOPProdSai_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialOP As New AdmFiliais
Dim objRastroLote As New ClassRastreamentoLote
Dim vbMsgRes As VbMsgBoxResult
Dim iEntradaSaida As Integer

On Error GoTo Erro_FilialOPProdSai_Validate

    If Len(Trim(FilialOPProdSai.Text)) <> 0 Then

        'Verifica se é uma FilialOP selecionada
        If FilialOPProdSai.Text <> FilialOPProdSai.List(FilialOPProdSai.ListIndex) Then

            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOPProdSai, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 126649

            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then

                objFilialOP.iCodFilial = iCodigo

                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 126650

                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 126651

                'coloca na tela
                FilialOPProdSai.Text = iCodigo & SEPARADOR & objFilialOP.sNome

            End If

            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 126652

        End If

        If Len(Trim(LoteProdSai.Text)) > 0 Then

                objRastroLote.sCodigo = LoteProdSai.Text
                objRastroLote.sProduto = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
                objRastroLote.iFilialOP = FilialOPProdSai.Text

                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 126653

                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 126654

                'Preenche a Quantidade do Lote
                lErro = QuantLote_Calcula(QuantDispProdSai)
                If lErro <> SUCESSO Then gError 126655


        End If

    Else

        'Preenche a Quantidade do Lote
        lErro = QuantDisponivel_Calcula1(QuantDispProdSai)
        If lErro <> SUCESSO Then gError 126656

    End If

    If Len(Trim(QuantDispProdSai.Caption)) <> 0 Then

        lErro = Testa_QuantRequisitada(QuantDispProdSai)
        If lErro <> SUCESSO Then gError 126657

    End If

    Exit Sub

Erro_FilialOPProdSai_Validate:

    Cancel = True

    Select Case gErr

        Case 126649, 126650, 126653, 126655, 126656, 126657

        Case 126651
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOPProdSai.Text)

        Case 126652
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOPProdSai.Text)

        Case 126654
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Chama_Tela_Modal("RastreamentoLote", objRastroLote)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174183)

    End Select

End Sub

Private Sub objEventoLocalizacao_evSelecao(obj1 As Object)

Dim objEstoqueProduto As ClassEstoqueProduto

    Set objEstoqueProduto = obj1

    'Se não tiver nenhuma linha do Grid selecionada --> Sai
    If Len(Trim(CodigoFilho.Caption)) = 0 Then Exit Sub
    
    'coloca o Almoxarifado no Grid na linha selecionada
    Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = objEstoqueProduto.iAlmoxarifado
    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    Me.Show

    Exit Sub

End Sub

Private Sub LabelVersao_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LabelVersao_Click

    objKit.sProdutoRaiz = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
    If Len(Trim(Versao.Text)) > 0 Then objKit.sVersao = Versao.Text
        
    colSelecao.Add objKit.sProdutoRaiz
    
    Call Chama_Tela_Modal("KitVersaoLista", colSelecao, objKit, objEventoVersao)
    
    Exit Sub

Erro_LabelVersao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174184)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long
Dim obj As ClassItemRomaneioGrade

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1

    Versao.Text = objKit.sVersao

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sVersao = objKit.sVersao

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174185)

    End Select

    Exit Sub
    
End Sub

Private Sub AlmoxOP_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AlmoxOP_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AlmoxOP_Validate

    lErro = AlmoxGeral_Validate(AlmoxOP)
    If lErro <> SUCESSO Then gError 126593

    Exit Sub

Erro_AlmoxOP_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 126593

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174186)

    End Select

    Exit Sub

End Sub

Private Sub LabelAlmoxOP_Click()
'Informa se produto é estocado em algum almoxarifado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelAlmoxOP_Click

    colSelecao.Add gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
    
    'chama a tela de lista de estoque do produto corrente
    Call Chama_Tela_Modal("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)

    
    Exit Sub

Erro_LabelAlmoxOP_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174187)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_objEventoEstoque_evselecao

    Set objEstoqueProduto = obj1

    AlmoxOP.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = objEstoqueProduto.iAlmoxarifado
    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = objEstoqueProduto.sAlmoxarifadoNomeReduzido

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174188)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is AlmoxOP Then
            Call LabelAlmoxOP_Click
        ElseIf Me.ActiveControl Is Versao Then
            Call LabelVersao_Click
        ElseIf Me.ActiveControl Is AlmoxProdSai Then
            Call LabelAlmoxProdSai_Click
        ElseIf Me.ActiveControl Is CodOPProdSai Then
            Call LabelOPProdSai_Click
        ElseIf Me.ActiveControl Is ProdOPProdSai Then
            Call LabelProdOPProdSai_Click
        ElseIf Me.ActiveControl Is CodOPProd Then
            Call LabelOPProd_Click
        ElseIf Me.ActiveControl Is AlmoxProd Then
            Call LabelAlmoxProd_Click
        End If
        
    End If
    
    If KeyCode <> vbKeyControl And Shift = 2 Then
        Call Trata_Tecla_Grid(KeyCode)
    End If

End Sub

Private Sub AlmoxProdSai_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AlmoxProdSai_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AlmoxProdSai_Validate

    lErro = AlmoxGeral_Validate(AlmoxProdSai)
    If lErro <> SUCESSO Then gError 126594

    'Se o Almoxarifado está preenchido
    If Len(Trim(AlmoxProdSai.Text)) > 0 Then
    
        If Len(Trim(gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sLote)) = 0 Then
            lErro = QuantDisponivel_Calcula1(QuantDispProdSai)
            If lErro <> SUCESSO Then gError 126637
        Else
            lErro = QuantLote_Calcula(QuantDispProdSai)
            If lErro <> SUCESSO Then gError 126638
        End If

        If Len(Trim(QuantDispProdSai.Caption)) <> 0 Then
    
            lErro = Testa_QuantRequisitada(QuantDispProdSai)
            If lErro <> SUCESSO Then gError 126661
    
        End If

    End If

    Exit Sub

Erro_AlmoxProdSai_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 126594, 126637, 126638, 126661

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174189)

    End Select

    Exit Sub

End Sub

Private Sub LabelAlmoxProdSai_Click()
'Informa se produto é estocado em algum almoxarifado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelAlmoxProdSai_Click

    colSelecao.Add gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
    
    'chama a tela de lista de estoque do produto corrente
    Call Chama_Tela_Modal("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoqueProdSai)
    
    Exit Sub

Erro_LabelAlmoxProdSai_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174190)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoqueProdSai_evselecao(obj1 As Object)

Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lErro As Long

On Error GoTo Erro_objEventoEstoqueProdSai_evselecao

    Set objEstoqueProduto = obj1

    AlmoxProdSai.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = objEstoqueProduto.iAlmoxarifado
    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    If Len(Trim(gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sLote)) = 0 Then
    
        'Calcula a Quantidade Disponível nesse Almoxarifado
         lErro = QuantDisponivel_Calcula1(QuantDispProdSai)
        If lErro <> SUCESSO Then gError 126639
    Else
        lErro = QuantLote_Calcula(QuantDispProdSai)
        If lErro <> SUCESSO Then gError 126640
    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoqueProdSai_evselecao:

    Select Case gErr

        Case 126639, 126640

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174191)

    End Select

    Exit Sub

End Sub

Private Function QuantDisponivel_Calcula1(ByVal objQuantDisp As Object) As Long
'descobre a quantidade disponivel e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objItem As ClassItemRomaneioGrade
Dim objProduto As New ClassProduto

On Error GoTo Erro_QuantDisponivel_Calcula1

    objQuantDisp.Caption = ""

    If gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado <> 0 Then
    
        objProduto.sCodigo = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto

        objEstoqueProduto.iAlmoxarifado = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado
        objEstoqueProduto.sProduto = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
    
        'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 126614
    
        'Se não encontrou EstoqueProduto no Banco de Dados
        If lErro = 21306 Then
        
             objQuantDisp.Caption = Formata_Estoque(0)

        Else
            
            If gobjRomaneioGrade.objObjetoTela.iBenef = MARCADO Then
            
                objQuantDisp.Caption = Formata_Estoque(objEstoqueProduto.dQuantBenef3)
        
            Else

                objQuantDisp.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel)
            
            End If
            
        End If

    Else

        'Limpa a Quantidade Disponível da Tela
        objQuantDisp.Caption = ""

    End If
    
    QuantDisponivel_Calcula1 = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula1:

    QuantDisponivel_Calcula1 = gErr

    Select Case gErr

        Case 126610, 126614, 126615

        Case 126611
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174192)

    End Select

    Exit Function

End Function

Private Function Testa_QuantRequisitada(objQuantDisp As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoAtual As String
Dim sAlmoxarifado As String
Dim sAlmoxarifadoAtual As String
Dim sUnidadeAtual As String
Dim sUnidadeProd As String
Dim dQuantidadeProd As String
Dim dFator As Double
Dim objProduto As New ClassProduto, sLoteAtual As String, sLote As String
Dim dQuantTotal As Double, iFilialOPAtual As Integer, iFilialOP As Integer
Dim vbMsg As VbMsgBoxResult
Dim objControle As Control
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_Testa_QuantRequisitada

    sProdutoAtual = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
    sAlmoxarifadoAtual = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado
    sUnidadeAtual = gobjRomaneioGrade.objObjetoTela.sSiglaUMEst
    iFilialOPAtual = Codigo_Extrai(FilialOPProdSai.Text)
    sLoteAtual = LoteProdSai.Text
    dQuantTotal = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).dQuantidade

    If Len(sProdutoAtual) > 0 And Len(sAlmoxarifadoAtual) > 0 And Len(sUnidadeAtual) > 0 Then

        objProduto.sCodigo = sProdutoAtual

        'Lê o produto para saber qual é a sua ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 126620
    
        If lErro = 28030 Then gError 126621
    
        For iIndice = 1 To gobjRomaneioGrade.objTela.objGrid.iLinhasExistentes
    
            'Não pode somar a Linha atual
            If gobjRomaneioGrade.objTela.objGrid.objGrid.Row <> iIndice Then
    
                'se nao for um "pai" de grade
                If left(gobjRomaneioGrade.objTela.objGrid.objGrid.TextMatrix(iIndice, 0), 1) <> "#" Then
    
                    sCodProduto = gobjRomaneioGrade.objTela.objGrid.objGrid.TextMatrix(iIndice, gobjRomaneioGrade.objTela.iGrid_Produto_Col)
                    sAlmoxarifado = gobjRomaneioGrade.objTela.objGrid.objGrid.TextMatrix(iIndice, gobjRomaneioGrade.objTela.iGrid_Almoxarifado_Col)
                    iFilialOP = Codigo_Extrai(gobjRomaneioGrade.objTela.objGrid.objGrid.TextMatrix(iIndice, gobjRomaneioGrade.objTela.iGrid_FilialOP_Col))
                    sLote = gobjRomaneioGrade.objTela.objGrid.objGrid.TextMatrix(iIndice, gobjRomaneioGrade.objTela.iGrid_Lote_Col)
                    dQuantidadeProd = StrParaDbl(gobjRomaneioGrade.objTela.objGrid.objGrid.TextMatrix(iIndice, gobjRomaneioGrade.objTela.iGrid_Quantidade_Col))
                    sUnidadeProd = gobjRomaneioGrade.objTela.objGrid.objGrid.TextMatrix(iIndice, gobjRomaneioGrade.objTela.iGrid_UnidadeMed_Col)
    
                    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 126622
        
                    'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                    If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) And iFilialOPAtual = iFilialOP And UCase(sLoteAtual) = UCase(sLote) Then
        
                        'Verifica se há alguma QuanTidade informada
                        If dQuantidadeProd <> 0 Then
        
                            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                            If lErro <> SUCESSO Then gError 126623
        
                            dQuantTotal = dQuantTotal + (dQuantidadeProd * dFator)
        
                        End If
                    
                    End If
        
                Else
                    'se for um "pai" de grade ==> pesquisa os filhos
                    Set objItemMovEst = gobjRomaneioGrade.objTela.gobjMovEst.colItens(iIndice)
                    
                    For Each objItemRomaneioGrade In objItemMovEst.colItensRomaneioGrade

                        'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                        If UCase(objItemRomaneioGrade.sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objItemRomaneioGrade.sProduto) = UCase(objProduto.sCodigo) And objItemRomaneioGrade.iFilialOP = iFilialOPAtual And UCase(objItemRomaneioGrade.sLote) = UCase(sLoteAtual) Then
            
                            'Verifica se há alguma QuanTidade informada
                            If objItemRomaneioGrade.dQuantidade <> 0 Then
            
                                dQuantTotal = dQuantTotal + objItemRomaneioGrade.dQuantidade
            
                            End If
                        
                        End If
                    
                    Next
                
                End If
        
            End If
    
        Next
    
        If dQuantTotal > StrParaDbl(objQuantDisp.Caption) Then
            vbMsg = Rotina_Aviso(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", gErr)
        End If

    End If

    Testa_QuantRequisitada = SUCESSO

    Exit Function

Erro_Testa_QuantRequisitada:

    Testa_QuantRequisitada = gErr

    Select Case gErr

        Case 126619, 126620, 126622, 126623

        Case 126621
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174193)

    End Select

    Exit Function

End Function

Private Function QuantLote_Calcula(objQuantDisp As Object) As Long
'descobre a quantidade Lote e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastreamentoLoteSaldo As New ClassRastreamentoLoteSaldo

On Error GoTo Erro_QuantLote_Calcula

    objQuantDisp.Caption = ""

    If Len(Trim(gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado)) <> 0 Then

        objRastreamentoLoteSaldo.iAlmoxarifado = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado
        objRastreamentoLoteSaldo.sProduto = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
        objRastreamentoLoteSaldo.sLote = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sLote
        objRastreamentoLoteSaldo.iFilialOP = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iFilialOP
        
        'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
        lErro = CF("RastreamentoLoteSaldo_Le", objRastreamentoLoteSaldo)
        If lErro <> SUCESSO And lErro <> 78633 Then gError 126635

        'Se não encontrou EstoqueProduto no Banco de Dados
        If lErro = 78633 Then
        
             objQuantDisp.Caption = Formata_Estoque(0)

        Else
    
            If gobjRomaneioGrade.objObjetoTela.iBenef = MARCADO Then
    
                objQuantDisp.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantBenef3)
                
            Else
    
                objQuantDisp.Caption = Formata_Estoque(objRastreamentoLoteSaldo.dQuantDispNossa)

            End If

        End If

    Else

        'Limpa a Quantidade Disponível da Tela
        objQuantDisp.Caption = ""

    End If
    
    QuantLote_Calcula = SUCESSO

    Exit Function

Erro_QuantLote_Calcula:

    QuantLote_Calcula = gErr

    Select Case gErr

        Case 126635

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174194)

    End Select

    Exit Function

End Function

Function AlmoxGeral_Validate(objAlmox As Object) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_AlmoxGeral_Validate

    'Se o Almoxarifado está preenchido
    If Len(Trim(objAlmox.Text)) > 0 Then

        'Valida o ALmoxarifado
        lErro = TP_Almoxarifado_Filial_Produto_Grid(gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto, objAlmox, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 126519
        'Se não for encontrado --> Erro
        If lErro = 25157 Then gError 126517
        If lErro = 25162 Then gError 126518

        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = objAlmoxarifado.iCodigo
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = objAlmoxarifado.sNomeReduzido

    Else
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = 0
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = ""
    
    End If

    AlmoxGeral_Validate = SUCESSO

    Exit Function

Erro_AlmoxGeral_Validate:
    
    AlmoxGeral_Validate = gErr

    Select Case gErr
    
        Case 126517
            'Pergunta de deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", objAlmox.Text)
            'Se a resposta for sim
            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = objAlmox.Text

                'Chama a Tela Almoxarifados
                Call Chama_Tela_Modal("Almoxarifado", objAlmoxarifado)

            End If

        Case 126518

            'Pergunta se deseja criar o Almoxarifado
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", Codigo_Extrai(objAlmox.Text))

            'Se a resposta for positiva
            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = Codigo_Extrai(objAlmox.Text)

                'Chama a tela de Almoxarifados
                Call Chama_Tela_Modal("Almoxarifado", objAlmoxarifado)

            End If
            
        Case 126519

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174195)

    End Select

    Exit Function

End Function

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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

Public Sub Form_Load()

    'Define o formato da quantidade
    Quantidade(0).Format = FORMATO_ESTOQUE
    
    Set objEventoLocalizacao = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoEstoqueProdSai = New AdmEvento
    Set objEventoOP = New AdmEvento
    Set objEventoProdutoOP = New AdmEvento
    Set objEventoOPProd = New AdmEvento
    Set objEventoEstoqueProd = New AdmEvento

    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort
    
    'Sinaliza que o Form_Loas ocorreu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

End Sub

Function Trata_Parametros(objRomaneioGrade As ClassRomaneioGrade) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objGrade As New ClassGrade
Dim sProdutoMascarado As String

On Error GoTo Erro_Trata_Parametros

    'Faz a variável global a tela apontar para a variável passada
    Set gobjRomaneioGrade = objRomaneioGrade
    
    'Verifica qual foi a tela que chamou a Grade
    Select Case gobjRomaneioGrade.sNomeTela
        
        'Caso seja a de Pedido de Venda, Orçamento
        Case NOME_TELA_PEDIDOVENDA, NOME_TELA_PEDIDOVENDACONSULTA
            'Deve receber o Funcionamento de Pedido
            gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_PEDIDO
    
            DescricaoPai.Caption = objRomaneioGrade.objObjetoTela.sDescricao
            
        Case NOME_TELA_ORCAMENTOVENDA
            'Deve receber o Funcionamento de Orcamento
            gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_ORCAMENTO
    
            DescricaoPai.Caption = objRomaneioGrade.objObjetoTela.sDescricao
        
        Case NOME_TELA_NFISCALFATURAPEDIDO, NOME_TELA_NFISCALPEDIDO
            'Deve receber o Funcionamento de Nota Fiscal Fatura Pedido
            gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_NFFATPEDIDO
            
            DescricaoPai.Caption = objRomaneioGrade.objObjetoTela.sDescricaoItem

        Case NOME_TELA_NFISCAL, NOME_TELA_NFISCALFATURA
            'Deve receber o Funcionamento de Nota Fiscal Fatura Pedido
            gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_NFISCAL
            
            DescricaoPai.Caption = objRomaneioGrade.objObjetoTela.sDescricaoItem

        Case NOME_TELA_RECEBMATERIALC, NOME_TELA_RECEBMATERIALF, NOME_TELA_NFISCALENTRADA, NOME_TELA_NFISCALFATENTRADA, NOME_TELA_NFISCALENTREM, NOME_TELA_NFISCALENTDEV, NOME_TELA_NFISCALREM, NOME_TELA_NFISCALDEV
        
            'Deve receber o Funcionamento de Nota Fiscal Fatura Pedido
            gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_RECEBIMENTO
            
            DescricaoPai.Caption = objRomaneioGrade.objObjetoTela.sDescricaoItem
       
        Case NOME_TELA_ORDEMPRODUCAO
            'Deve receber o Funcionamento de OP
            gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_OP
    
            DescricaoPai.Caption = objRomaneioGrade.objObjetoTela.sDescricao
       
        Case NOME_TELA_PRODUCAOSAIDA
            'Deve receber o Funcionamento de ProducaoSaida
            gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_PRODSAI
    
            DescricaoPai.Caption = objRomaneioGrade.objObjetoTela.sDescricao
       
            'Inicializa Máscara de Produto
            lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdOPProdSai)
            If lErro <> SUCESSO Then gError 126662
       
            'Carrega a combo de Filial O.P.
            lErro = Carrega_FilialOP()
            If lErro <> SUCESSO Then gError 126644
       
        Case NOME_TELA_PRODUCAOENTRADA
            'Deve receber o Funcionamento de OP
            gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_PRODENT
    
            DescricaoPai.Caption = objRomaneioGrade.objObjetoTela.sDescricao
       
    End Select

    'Passa o Produto para o obj
    objProduto.sCodigo = gobjRomaneioGrade.objObjetoTela.sProduto
            
    lErro = CF("Customiza_RomaneioGrade_TrataParam", objProduto)
    If lErro <> SUCESSO Then gError 117691
            
    'Lê o Produto passado
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 86297
    'Se o produto não existir ==> Erro
    If lErro <> SUCESSO Then gError 86298
    
    lErro = Mascara_MascararProduto(gobjRomaneioGrade.objObjetoTela.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 86311
        
    'Coloca o Produto Pai na Tela
    CodigoPai.Caption = sProdutoMascarado
    
    If gobjRomaneioGrade.sNomeTela = NOME_TELA_ORDEMPRODUCAO Then
        UnidadeMed.Caption = objRomaneioGrade.objObjetoTela.sSiglaUMEstoque
    ElseIf gobjRomaneioGrade.sNomeTela = NOME_TELA_PRODUCAOSAIDA Or gobjRomaneioGrade.sNomeTela = NOME_TELA_PRODUCAOENTRADA Then
        UnidadeMed.Caption = objRomaneioGrade.objObjetoTela.sSiglaUMEst
    Else
        UnidadeMed.Caption = objRomaneioGrade.objObjetoTela.sUnidadeMed
    End If
    
    'Inicializa a coleção que va guardar os
    'produtos de grade filhos do passado
    Set gobjRomaneioGrade.colItensRomaneioGrade = New Collection

    'Lê os filhos analíticos do produto pai de grade passado
    lErro = CF("Produto_Le_Filhos_Grade", objProduto, gobjRomaneioGrade.colItensRomaneioGrade)
    If lErro <> SUCESSO And lErro <> 86304 Then gError 86306
    
    'O produto não tem filhos de grade ou seus filhos são analíticos
    If lErro = 86304 Then gError 86307
    
    'Busca os dados das Categorias da Grade do Produto
    objGrade.sCodigo = objProduto.sGrade
    
    lErro = CF("GradeCategoria_Le", objGrade)
    If lErro <> SUCESSO Then gError 86309
    
    'Transfere para o objGlobal da tela as informações vindas da tela
    lErro = Transfere_Dados_Tela()
    If lErro <> SUCESSO Then gError 86312
    
    'Inicializa o Grid de Grade com as categorias da Grade
    'para as quais existem produtos filhos
    lErro = Inicializa_Grid_Grade(objGridGrade, objGrade)
    If lErro <> SUCESSO Then gError 86308

'    'Transfere para o objGlobal da tela as informações vindas da tela
'    lErro = Transfere_Dados_Tela()
'    If lErro <> SUCESSO Then gError 86312
    
    'Preenche o grid de Grade com as informações dos procutos filhos
    lErro = Preenche_Grid_Grade()
    If lErro <> SUCESSO Then gError 86313
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 86297, 86306, 86308, 86309, 86311, 117691, 126644, 126662
        
        Case 86298
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case 86307
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_FILHOS_GRADE", gErr, Trim(objProduto.sCodigo), Trim(objProduto.sGrade))
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174196)
    
    End Select
    
    Exit Function
        
End Function

Private Function Inicializa_Grid_Grade(objGridGrade As AdmGrid, objGrade As ClassGrade) As Long
'realiza a inicialização do grid

Dim objItemRomaneiro As ClassItemRomaneioGrade
Dim objGradeCateg As ClassGradeCategoria
Dim objCategProdItem As ClassCategoriaProdutoItem
Dim iIndice As Integer
Dim objGradeLinha As New ClassGradeLinCol
Dim objGradeColuna As New ClassGradeLinCol
Dim objGradeLinColCat As ClassGradeLinColCat
Dim objGradeLinColCatItem As ClassGradeLinColCatItem
Dim colItensCategProd As Collection
Dim objCategoriaProduto As ClassCategoriaProduto
Dim iNumLinhas As Integer
Dim iNumColunas As Integer
Dim bAchou As Boolean
Dim lErro As Long
Dim objCategProdItemAux As New ClassCategoriaProdutoItem
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objProd As ClassProduto

On Error GoTo Erro_Inicializa_Grid_Grade

    Set objGridGrade = New AdmGrid

    'tela em questão
    Set objGridGrade.objForm = Me

    'Para cada categoria da Grade
    For Each objGradeCateg In objGrade.colCategoria
        
        Set colItensCategProd = New Collection
        Set objCategoriaProduto = New ClassCategoriaProduto

        'Carrega a informação da Categoria no obj de leitura
        objCategoriaProduto.sCategoria = objGradeCateg.sCategoria

        'Busca no BD os Itens (Ordenados) da categoria passada
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategProd)
        If lErro <> SUCESSO And lErro <> 22541 Then gError ERRO_SEM_MENSAGEM
        If lErro = 22541 Then gError 86384
        
        Set objGradeLinColCat = New ClassGradeLinColCat
        Set objGradeLinColCat.objGradeCategoria = objGradeCateg
        For Each objCategProdItem In colItensCategProd
            Set objGradeLinColCatItem = New ClassGradeLinColCatItem
            Set objGradeLinColCatItem.objCategoriaProdutoItem = objCategProdItem
            
            bAchou = False
            
            For Each objItemRomaneiro In gobjRomaneioGrade.colItensRomaneioGrade
                Set objProd = New ClassProduto
                objProd.sCodigo = objItemRomaneiro.sProduto
                lErro = CF("Produto_Le", objProd)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                If objProd.iAtivo = PRODUTO_ATIVO Or objItemRomaneiro.dQuantidade > 0 Then
                    'Percorre as suas categorias
                    For Each objCategProdItemAux In objItemRomaneiro.colCategoria
                        If UCase(objCategProdItemAux.sCategoria) = UCase(objGradeCateg.sCategoria) Then
                            If UCase(objCategProdItemAux.sItem) = UCase(objCategProdItem.sItem) Then
                                bAchou = True
                                Exit For
                            End If
                        End If
                    Next
                    If bAchou Then Exit For
                End If
            Next
            
            If bAchou Then objGradeLinColCat.colItens.Add objGradeLinColCatItem

        Next
        
        If objGradeLinColCat.colItens.Count > 0 Then
            'Se for a categoria que é Linha
            If objGradeCateg.iPosicao = 0 Then
                objGradeLinha.colGradeCategorias.Add objGradeLinColCat
            'Se for a categoria que deve ocupar a posição das colunas
            Else
                objGradeColuna.colGradeCategorias.Add objGradeLinColCat
            End If
        End If
    
    Next
   
    Call Grid_Monta_Colunas(objGradeColuna, objGridGrade, objGradeLinha.colGradeCategorias.Count)
        
    'Inclui uma coluna para totalizar as quantidades de cada Linha
    objGridGrade.colColuna.Add ("Total Linha")

    Call Grid_Monta_Linhas(objGradeLinha, objGridGrade, objGradeColuna.iColunas)

    objGridGrade.objGrid = GridGrade
    
    objGridGrade.iLinhasExistentes = objGradeLinha.iLinhas + 1

    'todas as linhas do grid
    objGridGrade.objGrid.Rows = objGradeLinha.iLinhas + 1 + objGradeColuna.colGradeCategorias.Count

    objGridGrade.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'linhas visiveis do grid
    objGridGrade.iLinhasVisiveis = IIf(objGradeLinha.iLinhas + 1 < 8, objGradeLinha.iLinhas + 1, 8)

    'Largura da primeira coluna
    GridGrade.ColWidth(0) = Quantidade(0).Width

    'Largura automática para as outras colunas
    objGridGrade.iGridLargAuto = GRID_LARGURA_MANUAL

    'Proibido incluir no grid
    objGridGrade.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Proibido excluir no grid
    objGridGrade.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGridGrade)
    
    Call Grid_Exibe_Colunas(objGradeColuna, objGradeLinha.colGradeCategorias.Count)

    Call Grid_Exibe_Linhas(objGradeLinha, objGradeColuna.colGradeCategorias.Count)
    
    Set gobjGradeColuna = objGradeColuna
    Set gobjGradeLinha = objGradeLinha
    
    'Cria uma Linha para totalizar as colunas
    GridGrade.TextMatrix(objGradeLinha.iLinhas + objGradeColuna.colGradeCategorias.Count, 0) = ("Total Colunas")
    
    'Reposiciona os controles
    LabelProdutoGrade.top = GridGrade.top + GridGrade.Height + 200
    CodigoFilho.top = LabelProdutoGrade.top
    LabelDescricaoGrade.top = (LabelProdutoGrade.top + LabelProdutoGrade.Height) + 200
    DescricaoFilho.top = LabelDescricaoGrade.top
    Figura.top = LabelProdutoGrade.top - 100
                   
    'Verifica qual foi a tela que chamou a Grade
    Select Case gobjRomaneioGrade.iModoFuncionamento
                   
        Case ROMANEIOGRADE_FUNCIONAMENTO_ORCAMENTO, ROMANEIOGRADE_FUNCIONAMENTO_NFFATPEDIDO, ROMANEIOGRADE_FUNCIONAMENTO_RECEBIMENTO
            
            'Esconde o frame de quantidades
            FrameQuantidades.Visible = False
            FrameAlmoxarifado.Visible = False
            FrameOP.Visible = False
            FrameProdSai.Visible = False
            FrameProd.Visible = False

            FrameGrade.Height = GridGrade.Height + 1600

        Case ROMANEIOGRADE_FUNCIONAMENTO_NFISCAL

            'Esconde o frame de quantidades
            FrameQuantidades.Visible = False
            FrameOP.Visible = False
            FrameProdSai.Visible = False
            FrameProd.Visible = False
            
            FrameGrade.Height = GridGrade.Height + 1600
                        
        Case ROMANEIOGRADE_FUNCIONAMENTO_PEDIDO

            FrameOP.Visible = False
            FrameProdSai.Visible = False
            FrameProd.Visible = False
            
            FrameQuantidades.top = GridGrade.Height + 1600

            FrameGrade.Height = FrameQuantidades.top + FrameQuantidades.Height + 200

        Case ROMANEIOGRADE_FUNCIONAMENTO_NFISCALREM
            
            'Esconde o frame de quantidades
            FrameQuantidades.Visible = False
            FrameOP.Visible = False
            FrameProdSai.Visible = False
            FrameProd.Visible = False
                                   
            FrameAlmoxarifado.top = GridGrade.Height + 1600
            
            FrameGrade.Height = FrameAlmoxarifado.top + FrameAlmoxarifado.Height + 200
        
        Case ROMANEIOGRADE_FUNCIONAMENTO_OP
            
            'Esconde o frame de quantidades
            FrameQuantidades.Visible = False
            FrameAlmoxarifado.Visible = False
            FrameProdSai.Visible = False
            FrameProd.Visible = False
            
            FrameOP.Visible = True

            FrameOP.top = GridGrade.Height + 1600
            
            FrameGrade.Height = FrameOP.top + FrameOP.Height + 200
        
        Case ROMANEIOGRADE_FUNCIONAMENTO_PRODSAI
            
            'Esconde o frame de quantidades
            FrameQuantidades.Visible = False
            FrameAlmoxarifado.Visible = False
            FrameOP.Visible = False
            FrameProd.Visible = False
            
            FrameProdSai.Visible = True

            FrameProdSai.top = GridGrade.Height + 1600
            
            FrameGrade.Height = FrameProdSai.top + FrameProdSai.Height + 200
        
        Case ROMANEIOGRADE_FUNCIONAMENTO_PRODENT
            
            'Esconde o frame de quantidades
            FrameQuantidades.Visible = False
            FrameAlmoxarifado.Visible = False
            FrameOP.Visible = False
            FrameProdSai.Visible = False

            FrameProd.Visible = True
            
            FrameProd.top = GridGrade.Height + 1600
            
            FrameGrade.Height = FrameProd.top + FrameProd.Height + 200
        
        
    End Select

    BotaoOK.top = FrameGrade.top + FrameGrade.Height + 100
    
    BotaoCancela.top = BotaoOK.top
   
    UserControl.Parent.Height = (BotaoOK.top + BotaoOK.Height + 500)
    'Me.Height = (BotaoOK.top + BotaoOK.Height + 500)
        
    Inicializa_Grid_Grade = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Grade:

    Inicializa_Grid_Grade = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 86384
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174196)
    
    End Select
    
    Exit Function

End Function

Function Transfere_Dados_Tela() As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio existente do BD (Produtos Filhos do Produto passado)
    For Each objItemRomaneioGrade In gobjRomaneioGrade.colItensRomaneioGrade
        'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
        For Each objItemRomaneioGradeTela In gobjRomaneioGrade.objObjetoTela.colItensRomaneioGrade
            'Se encontrou o Item
            If UCase(objItemRomaneioGrade.sProduto) = UCase(objItemRomaneioGradeTela.sProduto) Then
                'Transfere as informações vindas da tela chamadora para essa tela
                objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
                objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
                objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
                objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
                objItemRomaneioGrade.dQuantFaturada = objItemRomaneioGradeTela.dQuantFaturada
                
                If objItemRomaneioGradeTela.dQuantidade <> 0 Then
                    objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
                Else
                    objItemRomaneioGrade.dQuantidade = -1
                End If
                
                objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
                objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
                objItemRomaneioGrade.dQuantCancelada = objItemRomaneioGradeTela.dQuantCancelada
                objItemRomaneioGrade.dQuantPV = objItemRomaneioGradeTela.dQuantPV
                objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntItemPV
                objItemRomaneioGrade.iAlmoxarifado = objItemRomaneioGradeTela.iAlmoxarifado
                objItemRomaneioGrade.sAlmoxarifado = objItemRomaneioGradeTela.sAlmoxarifado
                objItemRomaneioGrade.sVersao = objItemRomaneioGradeTela.sVersao
                objItemRomaneioGrade.sCodOP = objItemRomaneioGradeTela.sCodOP
                objItemRomaneioGrade.sProdOP = objItemRomaneioGradeTela.sProdOP
                objItemRomaneioGrade.sLote = objItemRomaneioGradeTela.sLote
                objItemRomaneioGrade.iFilialOP = objItemRomaneioGradeTela.iFilialOP
                objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
                
                'Transfere as informações de Localização
                Set objItemRomaneioGrade.colLocalizacao = New Collection
                
                For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
                    
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
                    objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
                    objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
                    objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
                    objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
                    objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
                    objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                    
                    objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                    
                Next
                                
            End If
        
        Next
    Next

    Exit Function

End Function

Function Preenche_Grid_Grade() As Long

Dim iNumLinhas As Integer
Dim iNumColunas As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade

    'Para cada Coluna de quantidade da tela
    For iNumColunas = gobjGradeLinha.colGradeCategorias.Count To gobjGradeColuna.iColunas + gobjGradeLinha.colGradeCategorias.Count - 1
        'Percorre as linhas dessa coluna
        For iNumLinhas = gobjGradeColuna.colGradeCategorias.Count To gobjGradeLinha.iLinhas + gobjGradeColuna.colGradeCategorias.Count - 1
            'Para cada item de romaneio possível
            For Each objItemRomaneio In gobjRomaneioGrade.colItensRomaneioGrade
            
                'Se o item corresponder ao da célula (Linha, coluna) em questão
                If Verifica_Romaneio_Grid(objItemRomaneio, iNumLinhas, iNumColunas) Then
                    'Se a quantidade estiver preenchida ==> coloca na tela
                    If objItemRomaneio.dQuantidade > 0 Then GridGrade.TextMatrix(iNumLinhas, iNumColunas) = Formata_Estoque(objItemRomaneio.dQuantidade)
                    If objItemRomaneio.dQuantidade = -1 Then GridGrade.TextMatrix(iNumLinhas, iNumColunas) = Formata_Estoque(0)
                    Exit For
                End If
                
            Next
        Next
    Next
    
    'A tualiza os totais de quantidades informadas
    Call Atualiza_Totais
    
    Exit Function
    
End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim sProdutoMascarado As String

On Error GoTo Erro_Rotina_Grid_Enable

    If iLocalChamada <> ROTINA_GRID_ABANDONA_CELULA Then

        objControl.Enabled = False
    
        'Limpa os dados particulares de cada célula
        CodigoFilho.Caption = ""
        DescricaoFilho.Caption = ""
        QuantCancelada.Text = ""
        QuantCancelada.Enabled = False
        QuantReservada.Caption = ""
        QuantFaturada.Caption = ""
        Almoxarifado.Text = ""
        Call VisualizarFigura("")
        
        'Para cada Item de Romaneio existente
        
        For Each objItemRomaneio In gobjRomaneioGrade.colItensRomaneioGrade
            
            iIndice = iIndice + 1
            
            'Se ele for o da célula em questão
            If Verifica_Romaneio_Grid(objItemRomaneio, iLinha, objControl.Index) Then
               
               'GUarda que ele é o item atual
                gobjRomaneioGrade.iItemAtual = iIndice
                
                If gobjRomaneioGrade.sNomeTela <> NOME_TELA_PEDIDOVENDACONSULTA Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
                
                'Formata o Produto
                lErro = Mascara_MascararProduto(objItemRomaneio.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 86322
                
                'Informa na tela os dados do produto
                CodigoFilho.Caption = sProdutoMascarado
                DescricaoFilho.Caption = objItemRomaneio.sDescricao
                'Habilita o campo para informar a quantidade Cancelada
                QuantCancelada.Enabled = True
                
                Call VisualizarFigura(objItemRomaneio.sProduto)
                
                'Se a quantidade cancelada estiver preenchida ==> COloca na tela
                If objItemRomaneio.dQuantCancelada > 0 Then QuantCancelada.Text = Formata_Estoque(objItemRomaneio.dQuantCancelada)
                'Preenche a tela com os dados do item
                QuantReservada.Caption = Formata_Estoque(objItemRomaneio.dQuantReservada)
                QuantFaturada.Caption = Formata_Estoque(objItemRomaneio.dQuantFaturada)
                If objItemRomaneio.iAlmoxarifado > 0 Then
                
                    If Len(Trim(objItemRomaneio.sAlmoxarifado)) > 0 Then
                        Almoxarifado.Text = objItemRomaneio.sAlmoxarifado
                    Else
                        Almoxarifado.Text = objItemRomaneio.iAlmoxarifado
                        Call Almoxarifado_Validate(bSGECancelDummy)
                    End If
                
                End If
                
                'Verifica qual foi a tela que chamou a Grade
                Select Case gobjRomaneioGrade.iModoFuncionamento
                
                    Case ROMANEIOGRADE_FUNCIONAMENTO_OP
                
                        lErro = Rotina_Grid_Enable_OP(objItemRomaneio, objControl)
                        If lErro <> SUCESSO Then gError 126617
                            
                    Case ROMANEIOGRADE_FUNCIONAMENTO_PRODSAI
                
                        lErro = Rotina_Grid_Enable_ProdSai(objItemRomaneio, objControl)
                        If lErro <> SUCESSO Then gError 126618
                            
                    Case ROMANEIOGRADE_FUNCIONAMENTO_PRODENT
                
                        lErro = Rotina_Grid_Enable_ProdEnt(objItemRomaneio, objControl)
                        If lErro <> SUCESSO Then gError 126691
                            
                End Select
                
            End If
        
        Next

    End If

    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 86322, 126617, 126618, 126691
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174197)
            
    End Select

    Exit Sub

End Sub

Private Function Rotina_Grid_Enable_OP(ByVal objItemRomaneio As ClassItemRomaneioGrade, ByVal objControl As Object) As Long

Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
Dim iPadrao As Integer
Dim sAlmoxarifadoPadrao As String
Dim iAlmoxarifadoPadrao As Integer
Dim objItemOPGrade As New ClassItemOP
Dim iIndice1 As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Rotina_Grid_Enable_OP

    objKit.sProdutoRaiz = objItemRomaneio.sProduto
    
    Versao.Clear
    AlmoxOP.Text = ""
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 126515
    
    iPadrao = -1
    
    'Carrega a Combo com os Dados da Colecao
    For Each objKit In colKits
    
        Versao.AddItem (objKit.sVersao)
        
        'Se for a padrao -> Armazena
        If objKit.iSituacao = KIT_SITUACAO_PADRAO Then iPadrao = iIndice1
        
        iIndice1 = iIndice1 + 1
        
    Next

    If Len(objItemRomaneio.sVersao) > 0 Then Versao.Text = objItemRomaneio.sVersao
    If Len(objItemRomaneio.sAlmoxarifado) > 0 Then
        AlmoxOP.Text = objItemRomaneio.sAlmoxarifado
    
    ElseIf Len(Trim(gobjRomaneioGrade.objObjetoTela.sAlmoxarifadoNomeRed)) > 0 Then
        
        lErro = CF("EstoqueProduto_TestaAssociacao", objItemRomaneio.sProduto, gobjRomaneioGrade.objObjetoTela.sAlmoxarifadoNomeRed)
                
        If lErro = SUCESSO Then
        
            objAlmoxarifado.sNomeReduzido = gobjRomaneioGrade.objObjetoTela.sAlmoxarifadoNomeRed
        
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 126659
        
            If lErro = 25056 Then gError 126658
            
            AlmoxOP.Text = objAlmoxarifado.sNomeReduzido
            
            objItemRomaneio.iAlmoxarifado = objAlmoxarifado.iCodigo
            objItemRomaneio.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
        
        End If
    
    Else
    
        'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
        lErro = CF("AlmoxarifadoPadrao_Le_CodNomeRed", giFilialEmpresa, objItemRomaneio.sProduto, sAlmoxarifadoPadrao, iAlmoxarifadoPadrao)
        If lErro <> SUCESSO Then gError 126520
        
        'preenche o grid
        AlmoxOP.Text = sAlmoxarifadoPadrao
        objItemRomaneio.iAlmoxarifado = iAlmoxarifadoPadrao
        objItemRomaneio.sAlmoxarifado = sAlmoxarifadoPadrao
        
    End If

    If iPadrao <> -1 And Versao.ListIndex = -1 Then Versao.ListIndex = iPadrao

    If gobjRomaneioGrade.objObjetoTela.lNumIntDoc <> 0 Then

        objItemOPGrade.lNumItemOP = gobjRomaneioGrade.objObjetoTela.lNumIntDoc
        objItemOPGrade.sProduto = objItemRomaneio.sProduto

        'se o item já estiver cadastrado ==> impede a alteracao dos atributos
        lErro = CF("ItensOrdemProducao_Le_NumItemOPProduto", objItemOPGrade)
        If lErro <> SUCESSO And lErro <> 126568 Then gError 126560

    End If

    If lErro = SUCESSO And gobjRomaneioGrade.objObjetoTela.lNumIntDoc <> 0 Then
        objControl.Enabled = False
        Versao.Enabled = False
        AlmoxOP.Enabled = False
    Else
        objControl.Enabled = True
        Versao.Enabled = True
        AlmoxOP.Enabled = True
    End If
    
    Rotina_Grid_Enable_OP = SUCESSO
    
    Exit Function

Erro_Rotina_Grid_Enable_OP:

    Rotina_Grid_Enable_OP = gErr

    Select Case gErr
    
        Case 126515, 126520, 126560, 126659
        
        Case 126658
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174198)
            
    End Select

    Exit Function

End Function

Private Function Rotina_Grid_Enable_ProdSai(ByVal objItemRomaneio As ClassItemRomaneioGrade, ByVal objControl As Object) As Long

Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
Dim iPadrao As Integer
Dim sAlmoxarifadoPadrao As String
Dim iAlmoxarifadoPadrao As Integer
Dim objItemOPGrade As New ClassItemOP
Dim iIndice1 As Integer
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoEnxuto As String

On Error GoTo Erro_Rotina_Grid_Enable_ProdSai

    AlmoxProdSai.Text = ""
    CodOPProdSai.Text = ""
    ProdOPProdSai.PromptInclude = False
    ProdOPProdSai.Text = ""
    ProdOPProdSai.PromptInclude = True
    LoteProdSai.Text = ""
    FilialOPProdSai.ListIndex = -1
    
    If Len(Trim(objItemRomaneio.sAlmoxarifado)) > 0 Then
        
        AlmoxProdSai.Text = objItemRomaneio.sAlmoxarifado
    
    ElseIf Len(Trim(gobjRomaneioGrade.objObjetoTela.sAlmoxarifadoNomeRed)) <> 0 Then
        'se o almoxarifado padrao foi fornecido pelo usuario ==> tenta usa-lo
        lErro = CF("EstoqueProduto_TestaAssociacao", objItemRomaneio.sProduto, gobjRomaneioGrade.objObjetoTela.sAlmoxarifadoNomeRed)
        
        If lErro = SUCESSO Then
            
            objAlmoxarifado.sNomeReduzido = gobjRomaneioGrade.objObjetoTela.sAlmoxarifadoNomeRed
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 126645
    
            If lErro = 25056 Then gError 126646
            
            AlmoxProdSai.Text = objAlmoxarifado.sNomeReduzido
            
            objItemRomaneio.iAlmoxarifado = objAlmoxarifado.iCodigo
            objItemRomaneio.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
            
        End If
    
    Else
    
        'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
        lErro = CF("AlmoxarifadoPadrao_Le_CodNomeRed", giFilialEmpresa, objItemRomaneio.sProduto, sAlmoxarifadoPadrao, iAlmoxarifadoPadrao)
        If lErro <> SUCESSO Then gError 126647
        
        'preenche o grid
        AlmoxProdSai.Text = sAlmoxarifadoPadrao
        objItemRomaneio.iAlmoxarifado = iAlmoxarifadoPadrao
        objItemRomaneio.sAlmoxarifado = sAlmoxarifadoPadrao
        
    End If

    If Len(Trim(objItemRomaneio.sCodOP)) > 0 Then
        
        CodOPProdSai.Text = objItemRomaneio.sCodOP
    
    ElseIf Len(Trim(gobjRomaneioGrade.objObjetoTela.sOPCodigo)) > 0 Then
        
        objItemRomaneio.sCodOP = gobjRomaneioGrade.objObjetoTela.sOPCodigo
        
        CodOPProdSai.Text = objItemRomaneio.sCodOP
        
    End If

    If Len(Trim(objItemRomaneio.sProdOP)) > 0 Then
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemRomaneio.sProdOP, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 126663
    
        'Coloca o Codigo na tela
        ProdOPProdSai.PromptInclude = False
        ProdOPProdSai.Text = sProdutoEnxuto
        ProdOPProdSai.PromptInclude = True
    
    End If
    
    'se o item já estiver cadastrado
    If gobjRomaneioGrade.objObjetoTela.lNumIntDoc <> 0 Then
        objControl.Enabled = False
        AlmoxProdSai.Enabled = False
        CodOPProdSai.Enabled = False
        ProdOPProdSai.Enabled = False
        LoteProdSai.Enabled = False
    Else
        objControl.Enabled = True
        AlmoxProdSai.Enabled = True
        CodOPProdSai.Enabled = True
        ProdOPProdSai.Enabled = True
        LoteProdSai.Enabled = True
    End If
    
    objProduto.sCodigo = objItemRomaneio.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 126643

    If lErro = 28030 Then gError 126644
    
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
    
        FilialOPProdSai.Enabled = True
        
        LoteProdSai.Text = objItemRomaneio.sLote
        
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
        
        FilialOPProdSai.Enabled = True
        
        LoteProdSai.Text = objItemRomaneio.sLote
        
        If objItemRomaneio.iFilialOP <> 0 Then
        
            For iIndice = 0 To FilialOPProdSai.ListCount - 1
                If FilialOPProdSai.ItemData(iIndice) = objItemRomaneio.iFilialOP Then
                    FilialOPProdSai.ListIndex = iIndice
                    Exit For
                End If
            Next
        
        Else
        
            For iIndice = 0 To FilialOPProdSai.ListCount - 1
                If FilialOPProdSai.ItemData(iIndice) = giFilialEmpresa Then
                    FilialOPProdSai.ListIndex = iIndice
                    objItemRomaneio.iFilialOP = giFilialEmpresa
                    Exit For
                End If
            Next
        End If
    Else
        FilialOPProdSai.Enabled = False
    End If
    
    If Len(Trim(gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sLote)) = 0 Then
        lErro = QuantDisponivel_Calcula1(QuantDispProdSai)
        If lErro <> SUCESSO Then gError 126637
    Else
        lErro = QuantLote_Calcula(QuantDispProdSai)
        If lErro <> SUCESSO Then gError 126638
    End If
    
    Rotina_Grid_Enable_ProdSai = SUCESSO
    
    Exit Function

Erro_Rotina_Grid_Enable_ProdSai:

    Rotina_Grid_Enable_ProdSai = gErr

    Select Case gErr
    
        Case 126643, 126645, 126647, 126663
    
        Case 126644
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 126646
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174199)
            
    End Select

    Exit Function

End Function

Private Function Rotina_Grid_Enable_ProdEnt(ByVal objItemRomaneio As ClassItemRomaneioGrade, ByVal objControl As Object) As Long

Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
Dim iPadrao As Integer
Dim sAlmoxarifadoPadrao As String
Dim iAlmoxarifadoPadrao As Integer
Dim objItemOPGrade As New ClassItemOP
Dim iIndice1 As Integer
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoEnxuto As String

On Error GoTo Erro_Rotina_Grid_Enable_ProdEnt

    AlmoxProd.Text = ""
    CodOPProd.Text = ""
    LoteProd.Text = ""
    HorasMaqProd.Text = ""
    
    If Len(Trim(objItemRomaneio.sAlmoxarifado)) > 0 Then
        
        AlmoxProd.Text = objItemRomaneio.sAlmoxarifado
    
    Else
    
        'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
        lErro = CF("AlmoxarifadoPadrao_Le_CodNomeRed", giFilialEmpresa, objItemRomaneio.sProduto, sAlmoxarifadoPadrao, iAlmoxarifadoPadrao)
        If lErro <> SUCESSO Then gError 126690
        
        'preenche o grid
        AlmoxProd.Text = sAlmoxarifadoPadrao
        objItemRomaneio.iAlmoxarifado = iAlmoxarifadoPadrao
        objItemRomaneio.sAlmoxarifado = sAlmoxarifadoPadrao
        
    End If

    If Len(Trim(objItemRomaneio.sCodOP)) > 0 Then
        
        CodOPProd.Text = objItemRomaneio.sCodOP
    
    ElseIf Len(Trim(gobjRomaneioGrade.objObjetoTela.sOPCodigo)) > 0 Then
        
        objItemRomaneio.sCodOP = gobjRomaneioGrade.objObjetoTela.sOPCodigo
        
        CodOPProd.Text = objItemRomaneio.sCodOP
        
    End If

    If objItemRomaneio.lHorasMaquina > 0 Then
        HorasMaqProd.Text = objItemRomaneio.lHorasMaquina
    Else
        HorasMaqProd.Text = ""
    End If
    
    'se o item já estiver cadastrado
    If gobjRomaneioGrade.objObjetoTela.lNumIntDoc <> 0 Then
        objControl.Enabled = False
        AlmoxProd.Enabled = False
        CodOPProd.Enabled = False
        HorasMaqProd.Enabled = False
    Else
        objControl.Enabled = True
        AlmoxProd.Enabled = True
        CodOPProd.Enabled = True
        HorasMaqProd.Enabled = True
    End If
    
    objProduto.sCodigo = objItemRomaneio.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 126691

    If lErro = 28030 Then gError 126692
    
    LoteProd.Enabled = False
    
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
        
        LoteProdSai.Text = objItemRomaneio.sLote
        
        If gobjRomaneioGrade.objObjetoTela.lNumIntDoc = 0 Then LoteProd.Enabled = True
        
    End If
    
    Rotina_Grid_Enable_ProdEnt = SUCESSO
    
    Exit Function

Erro_Rotina_Grid_Enable_ProdEnt:

    Rotina_Grid_Enable_ProdEnt = gErr

    Select Case gErr
    
        Case 126690, 126691
    
        Case 126692
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174200)
            
    End Select

    Exit Function

End Function

Private Sub GridGrade_Click()

Dim objCampo As Object



Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridGrade, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
    
        Call Grid_Entrada_Celula(objGridGrade, iAlterado)
        
    End If
    
    Set objCampo = objGridGrade.objForm.Controls(objGridGrade.colCampo(objGridGrade.objGrid.Col))
    
    If TypeName(objCampo) = "Object" Then Set objCampo = objCampo(objGridGrade.colIndex(objGridGrade.objGrid.Col))
    
    Call objGridGrade.objForm.Rotina_Grid_Enable(objGridGrade.objGrid.Row, objCampo, ROTINA_GRID_ENTRADA_CELULA)
    If objCampo.Enabled = True Then objCampo.SetFocus
    

End Sub

Private Sub GridGrade_EnterCell()

     Call Grid_Entrada_Celula(objGridGrade, iAlterado)

End Sub

Private Sub GridGrade_GotFocus()

    Call Grid_Recebe_Foco(objGridGrade)

End Sub

Private Sub GridGrade_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridGrade)

        iAlterado = REGISTRO_ALTERADO

    Exit Sub

End Sub

Private Sub GridGrade_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridGrade, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
    
        Call Grid_Entrada_Celula(objGridGrade, iAlterado)
    End If

End Sub

Private Sub GridGrade_LeaveCell()

    Call Saida_Celula(objGridGrade)

End Sub
Private Sub GridGrade_RowColChange()

    
    Call Grid_RowColChange(objGridGrade)
'mario    Call Rotina_Grid_Enable(GridGrade.Row, Quantidade(GridGrade.Col), 1)
    
End Sub

Private Sub GridGrade_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridGrade)

End Sub

'mario Function Saida_Celula(objGridCategoria As AdmGrid) As Long
Function Saida_Celula(objGridGrade As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridGrade)

    If lErro = SUCESSO Then

        lErro = Saida_Celula_Quantidade(objGridGrade)
        If lErro <> SUCESSO Then gError 123221

'mario        lErro = Grid_Finaliza_Saida_Celula(objGridCategoria)
        lErro = Grid_Finaliza_Saida_Celula(objGridGrade)
        If lErro <> SUCESSO Then gError 123222

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 123221, 123222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174201)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim bQuantidadeIgual As Boolean
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim dQuantidadeFaturada As Double
Dim iIndice As Integer
Dim dPrecoUnitario As Double
Dim dQuantAnterior As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade(GridGrade.Col)
    
    If gobjRomaneioGrade.iTipoNFiscal <> 0 Then
        objTipoDocInfo.iCodigo = gobjRomaneioGrade.iTipoNFiscal
        
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 26665
    Else
        objTipoDocInfo.iComplementar = MARCADO
    End If

    'Inicializa dizendo que a quantidade não é igual
    bQuantidadeIgual = False

    'Se o controle tem alguma coisa escrita
    If Len(Quantidade(GridGrade.Col)) > 0 Then

        'Valida se é é um valor e positivo
        If objTipoDocInfo.iComplementar = DESMARCADO Then
            lErro = Valor_Positivo_Critica(Quantidade(GridGrade.Col).Text)
        Else
            lErro = Valor_NaoNegativo_Critica(Quantidade(GridGrade.Col).Text)
        End If
        'lErro = Valor_Positivo_Critica(Quantidade(GridGrade.Col).Text)
        If lErro <> SUCESSO Then gError 26665

        'Formata o valor para exibição
        Quantidade(GridGrade.Col).Text = Formata_Estoque(Quantidade(GridGrade.Col).Text)

    End If

    'Pega as demais quantidades da tela desse intem
    dQuantidade = StrParaDbl(Quantidade(GridGrade.Col).Text)
    
    'Atualiza a quantidade
    'Se está preenchido com 0 tem que tratar diferente pois pode ser NF de complemento
    'Vou forçar -1 para tratar na gravação
    If dQuantidade = 0 And Len(Quantidade(GridGrade.Col)) > 0 Then
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).dQuantidade = -1
    Else
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).dQuantidade = dQuantidade
    End If
    
    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridGrade.TextMatrix(GridGrade.Row, (GridGrade.Col)))
    
    'Se a quantidade for igua a anterior ==> Sinaliza A igualdade
    If dQuantAnterior = StrParaDbl(Quantidade(GridGrade.Col).Text) Then bQuantidadeIgual = True
    
    'EScreve no grid
    GridGrade.TextMatrix(GridGrade.Row, (GridGrade.Col)) = Formata_Estoque(dQuantidade)

    'Se a quantidade foi alterada
    If Not bQuantidadeIgual Then
        
        If gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_PEDIDO Then
        
            dQuantidadeCancelada = StrParaDbl(QuantCancelada.Text)
            dQuantidadeFaturada = StrParaDbl(QuantFaturada.Caption)
            If dQuantidadeCancelada > 0 And dQuantidade < dQuantidadeCancelada Then gError 26666
            If dQuantidadeFaturada > 0 And dQuantidade - dQuantidadeCancelada < dQuantidadeFaturada Then gError 26667
    
            'Se necessário refaz a reserva
            lErro = Reserva_Processa(dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada)
            If lErro <> SUCESSO Then gError 26831
        
        ElseIf gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_NFFATPEDIDO Or gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_NFISCAL Or gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_RECEBIMENTO Then
        
            If gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_NFFATPEDIDO And dQuantidade > gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).dQuantPV Then gError 46934
            
            lErro = Alocacao_Processa(dQuantidade)
            If lErro <> SUCESSO Then gError 46937
            
        ElseIf gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_OP Then
                
            'Coloca o Produto no Formato do Banco de Dados
            lErro = CF("Produto_Formata", CodigoFilho.Caption, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 126511
            
            objProdutoFilial.sProduto = sProdutoFormatado
            objProdutoFilial.iFilialEmpresa = giFilialEmpresa
            
            'Busca o Lote Mínino do Produto/FilialEmpresa
            lErro = CF("ProdutoFilial_Le", objProdutoFilial)
            If lErro <> SUCESSO And lErro <> 28261 Then gError 126512
                
            'preenche Verifica se Lote mínimo esta preenchdo (caso exista)
            If objProdutoFilial.dLoteMinimo > 0 Then
            
                If dQuantidade < objProdutoFilial.dLoteMinimo Then gError 126513
            
            End If
        
        ElseIf gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_PRODSAI Then

            lErro = Testa_QuantRequisitada(QuantDispProdSai)
            If lErro <> SUCESSO Then gError 126648
        
        End If

    
    End If
    
    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 59727
   
    'Atualiza as células de Totais
    Call Atualiza_Totais
   
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:
    
    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 26831, 46937
            GridGrade.TextMatrix(GridGrade.Row, GridGrade.Col) = Formata_Estoque(dQuantAnterior)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 26666
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26667
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_SUPERIOR", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
       
        Case 46934
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_FATURAR_MENOR", gErr, gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).dQuantPV)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 59727, 26665, 51037, 81525, 126511, 126512, 126648
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 126513
            Call Rotina_Erro(vbOKOnly, "ERRO_QDTPRODUTO_MENOR_LOTEMININO", gErr, objProdutoFilial.dLoteMinimo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174202)

    End Select

    Exit Function

End Function

Private Function Reserva_Processa(dQuantidade As Double, dQuantidadeCancelada As Double, dQuantidadeFaturada As Double) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objItemRomaneioGrade As ClassItemRomaneioGrade

On Error GoTo Erro_Reserva_Processa

    Set objItemRomaneioGrade = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual)

    'Se não houve Quant a Reservar zera Reservas.
    If dQuantidade = 0 Or dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 Then
        QuantReservada.Caption = Format(0, "Standard")
    End If


    objProduto.sCodigo = objItemRomaneioGrade.sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 26669

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 26668
    
    objItemRomaneioGrade.sUMEstoque = objProduto.sSiglaUMEstoque

    'Se controla Reserva e há quantidade a reservar gera reserva automático
    If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA And dQuantidade > 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada > 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
        
        'tenta reservar no Almoxarifado padrão
        lErro = ReservaAlmoxarifadoPadrao(objProduto, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada)
        If lErro <> SUCESSO Then gError 26679

    'Se não há quantidade a reservar
    ElseIf dQuantidade = 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
        
        'Elimina as reservas feitas anteriormente
        Set objItemRomaneioGrade.colLocalizacao = New Collection
        objItemRomaneioGrade.dQuantReservada = 0
                        
    End If

    Reserva_Processa = SUCESSO

    Exit Function

Erro_Reserva_Processa:

    Reserva_Processa = gErr

    Select Case gErr

        Case 26796, 26669, 26679

        Case 26668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174203)

    End Select

    Exit Function

End Function

Function ReservaAlmoxarifadoPadrao(objProduto As ClassProduto, dQuantidade As Double, dQuantCancelada As Double, dQuantFaturada As Double) As Long

Dim lErro As Long
Dim dQuantidadeReservarVenda As Double
Dim dQuantReservadaPedido As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim dFator As Double
Dim iAlmoxarifadoPadrao As Integer
Dim objItemPV As New ClassItemPedido
Dim objReservaBD As New ClassReserva
Dim objReservaItem As ClassReservaItem
Dim objReserva As ClassReserva
Dim colItemPedido As New colItemPedido
Dim objAlmoxarifadoPadrao As New ClassAlmoxarifado
Dim objAlmoxarifadoPadrao1 As New ClassAlmoxarifado
Dim iFilialEmpresa As Integer
Dim iFilialFaturamento As Integer
Dim dQuantidadeReservarEstoque1 As Double 'Reservar Filial 1
Dim dQuantidadeReservarEstoque2 As Double 'Reservar Filial 2
Dim objEstoqueProduto1 As New ClassEstoqueProduto 'Reservar Filial 1
Dim objEstoqueProduto2 As New ClassEstoqueProduto 'Reservar Filial 2

On Error GoTo Erro_ReservaAlmoxarifadoPadrao

    Set objItemRomaneioGrade = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual)
    
    'calula a quantidade reservada
    dQuantidadeReservarVenda = (dQuantidade - dQuantCancelada - dQuantFaturada)

    lErro = CF("UM_Conversao", objProduto.iClasseUM, gobjRomaneioGrade.objObjetoTela.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 26670

    'Converte a quantidade para a unidade de Estoque
    dQuantidadeReservarEstoque1 = dQuantidadeReservarVenda * dFator

    iFilialFaturamento = gobjRomaneioGrade.iFilialFaturamento

    If iFilialFaturamento = 0 Then
        'Busca a filial de faturamento
        lErro = CF("FilialFaturamento_Le", giFilialEmpresa, iFilialFaturamento)
        If lErro <> SUCESSO Then gError 30414
        If iFilialFaturamento = 0 Then iFilialFaturamento = giFilialEmpresa
    End If

    'Busca o almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", iFilialFaturamento, objProduto.sCodigo, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 51417
    If iAlmoxarifadoPadrao = 0 Then gError 51418

    objAlmoxarifadoPadrao.iCodigo = iAlmoxarifadoPadrao
    'Lê o almoxarifado padrão
    lErro = CF("Almoxarifado_Le", objAlmoxarifadoPadrao)
    If lErro = SUCESSO Then

        objEstoqueProduto1.sProduto = objProduto.sCodigo
        objEstoqueProduto1.iAlmoxarifado = iAlmoxarifadoPadrao
        objEstoqueProduto1.sAlmoxarifadoNomeReduzido = objAlmoxarifadoPadrao.sNomeReduzido

        'Lê o estoque do produto
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto1)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 26672
        If lErro = 21306 Then gError 26673

        objItemPV.lCodPedido = gobjRomaneioGrade.objObjetoTela.lCodPedido
        objItemPV.sProduto = objProduto.sCodigo
        objItemPV.iFilialEmpresa = giFilialEmpresa
        objItemPV.sProdutoDescricao = objProduto.sDescricao
        objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao

        'Lê as reservas do produto
        lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
        If lErro <> SUCESSO And lErro <> 26678 Then gError 26674

        dQuantReservadaPedido = objReservaBD.dQuantidade

        'Saldo enxergado por esse Pedido Venda
        objEstoqueProduto1.dSaldo = objEstoqueProduto1.dQuantDisponivel + dQuantReservadaPedido

        'Se quantidade a reservar for menor que a disponível coloca reserva no GridAlocacao
        If (dQuantidadeReservarEstoque1 - objEstoqueProduto1.dSaldo) < QTDE_ESTOQUE_DELTA Then
        
            QuantReservada.Caption = Formata_Estoque(dQuantidadeReservarVenda)
            
            objItemRomaneioGrade.dQuantReservada = dQuantidadeReservarVenda
            
            Set objItemRomaneioGrade.colLocalizacao = New Collection
            
            Set objReservaItem = New ClassReservaItem
            
            objReservaItem.iAlmoxarifado = objEstoqueProduto1.iAlmoxarifado
            objReservaItem.sAlmoxarifado = objEstoqueProduto1.sAlmoxarifadoNomeReduzido
            objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
            objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque1)
            objReservaItem.dtDataValidade = DATA_NULA
            objReservaItem.sResponsavel = RESERVA_AUTO_RESP

            objItemRomaneioGrade.colLocalizacao.Add objReservaItem

        Else

            lErro = CF("Retorna_Almoxarifado_Alternativo", iAlmoxarifadoPadrao, iFilialEmpresa)
            If lErro <> SUCESSO Then gError 105004

            If iAlmoxarifadoPadrao = 0 Then GoTo FaltaEstoque

            'Verifica quanto falta para reservar
            dQuantidadeReservarEstoque2 = dQuantidadeReservarEstoque1 - objEstoqueProduto1.dSaldo
            dQuantidadeReservarEstoque1 = objEstoqueProduto1.dSaldo

            'Busca na Filial 2
            objAlmoxarifadoPadrao1.iCodigo = iAlmoxarifadoPadrao

            'Lê o almoxarifado
            lErro = CF("Almoxarifado_Le", objAlmoxarifadoPadrao1)
            If lErro = SUCESSO Then

                objEstoqueProduto2.sProduto = objProduto.sCodigo
                objEstoqueProduto2.iAlmoxarifado = iAlmoxarifadoPadrao
                objEstoqueProduto2.sAlmoxarifadoNomeReduzido = objAlmoxarifadoPadrao1.sNomeReduzido

                'Lê o estoque do produto
                lErro = CF("EstoqueProduto_Le", objEstoqueProduto2)
                If lErro <> SUCESSO And lErro <> 21306 Then gError 26672
                If lErro = 21306 Then gError 26673

                objItemPV.lCodPedido = gobjRomaneioGrade.objObjetoTela.lCodPedido
                objItemPV.sProduto = objProduto.sCodigo
                objItemPV.iFilialEmpresa = iFilialEmpresa
                objItemPV.sProdutoDescricao = objProduto.sDescricao
                objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao

                'Lê as reservas do produto
                lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
                If lErro <> SUCESSO And lErro <> 26678 Then gError 26674

                dQuantReservadaPedido = objReservaBD.dQuantidade

                'Saldo enxergado por esse Pedido Venda
                objEstoqueProduto2.dSaldo = objEstoqueProduto2.dQuantDisponivel + dQuantReservadaPedido

                'Se quantidade a reservar for menor que a disponível coloca reserva no GridAlocacao
                If (dQuantidadeReservarEstoque2 - objEstoqueProduto2.dSaldo) < QTDE_ESTOQUE_DELTA Then

                    If dQuantidadeReservarEstoque1 > 0 Then
                        
                        Set objReservaItem = New ClassReservaItem
                        
                        'Parte reservada na filial 1
                        objReservaItem.iAlmoxarifado = objEstoqueProduto1.iAlmoxarifado
                        objReservaItem.sAlmoxarifado = objEstoqueProduto1.sAlmoxarifadoNomeReduzido
                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque1)
                        objReservaItem.sResponsavel = RESERVA_AUTO_RESP
                        objReservaItem.dtDataValidade = DATA_NULA
                        
                        objItemRomaneioGrade.colLocalizacao.Add objReservaItem

                    End If

                    Set objReservaItem = New ClassReservaItem
                    
                    'Parte Reservada na Filial 2
                        objReservaItem.iAlmoxarifado = objEstoqueProduto2.iAlmoxarifado
                        objReservaItem.sAlmoxarifado = objEstoqueProduto2.sAlmoxarifadoNomeReduzido
                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque2)
                        objReservaItem.sResponsavel = RESERVA_AUTO_RESP
                        objReservaItem.dtDataValidade = DATA_NULA
                        
                        objItemRomaneioGrade.colLocalizacao.Add objReservaItem

                'Caso contrário limpa as reservas desse ítem e chama tela de Falta de Estoque
                Else

FaltaEstoque:
                    Set objItemPV = New ClassItemPedido
                    objItemPV.dQuantidade = dQuantidade
                    objItemPV.dQuantReservada = 0
                    objItemPV.dQuantCancelada = dQuantCancelada
                    objItemPV.dQuantFaturada = dQuantFaturada
                    objItemPV.sProduto = objProduto.sCodigo
                    objItemPV.lCodPedido = gobjRomaneioGrade.objObjetoTela.lCodPedido
                    objItemPV.iItem = gobjRomaneioGrade.objObjetoTela.iItem
                    objItemPV.lNumIntDoc = gobjRomaneioGrade.objObjetoTela.lNumIntDoc
                    objItemPV.sUMEstoque = objProduto.sSiglaUMEstoque
                    objItemPV.sProdutoDescricao = objProduto.sDescricao
                    objItemPV.sUnidadeMed = gobjRomaneioGrade.objObjetoTela.sUnidadeMed
                    objItemPV.iClasseUM = objProduto.iClasseUM
                    objItemPV.iPossuiGrade = MARCADO
                    
                    'Chama tela de Falta de Estoque
                    lErro = Chama_Tela_Modal("FaltaEstoque", objItemPV, colItemPedido, dQuantidadeReservarEstoque1, objAlmoxarifadoPadrao, objEstoqueProduto1.dSaldo)

                    'Se retornar Cancela erro
                    If giRetornoTela = vbCancel Then gError 26680

                    'Limpa reservas desse ítem no GridAlocacao
                    Set objItemRomaneioGrade.colLocalizacao = New Collection
                    
                    If giRetornoTela = vbOK Then

                       'Se não substituiu o Produto
                        'Coloca QuantReservada e QuantCancelada no ítem do GridItens
                        QuantReservada.Caption = Formata_Estoque(objItemPV.dQuantReservada)
                        objItemRomaneioGrade.dQuantReservada = objItemPV.dQuantReservada
                                                
                        If objItemPV.dQuantCancelada > 0 Then
                            QuantCancelada.Text = Formata_Estoque(StrParaDbl(QuantCancelada.Text) + objItemPV.dQuantCancelada)
                        End If
                        
                        objItemRomaneioGrade.dQuantCancelada = objItemPV.dQuantCancelada

                        'Coloca reseravas desse ítem no GridAlocacao
                        For Each objReserva In objItemPV.ColReserva
                            Set objReservaItem = New ClassReservaItem
                            
                            objReservaItem.iAlmoxarifado = objReserva.iAlmoxarifado
                            objReservaItem.dQuantidade = objReserva.dQuantidade
                            objReservaItem.dtDataValidade = objReserva.dtDataValidade
                            objReservaItem.sAlmoxarifado = objReserva.sAlmoxarifado
                            objReservaItem.sResponsavel = objReserva.sResponsavel
                            
                            objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                            
                        Next
                    End If
                End If
            End If
        End If
    End If

    ReservaAlmoxarifadoPadrao = SUCESSO

    Exit Function

Erro_ReservaAlmoxarifadoPadrao:

    ReservaAlmoxarifadoPadrao = gErr

    Select Case gErr

        Case 26670, 26672, 26674, 26682, 26772, 30414, 105004

        Case 26681
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 26673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto1.sProduto, objEstoqueProduto1.iAlmoxarifado)

        Case 26680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESERVA_NAO_DECIDIDA", gErr, objProduto.sCodigo)

        Case 51418
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174204)

    End Select

    Exit Function

End Function

Sub Atualiza_Totais()

Dim dTotalLinhas As Double
Dim dTotalColunas As Double
Dim iLinha As Integer
Dim iLinha2 As Integer
Dim iColuna As Integer
Dim dTotalGeral As Double

    'Para cada Linha do Grid
    For iLinha = gobjGradeColuna.colGradeCategorias.Count To gobjGradeLinha.iLinhas + gobjGradeColuna.colGradeCategorias.Count - 1

        'Em cada Linha ==> Percorre todas as colunas
        For iColuna = gobjGradeLinha.colGradeCategorias.Count To gobjGradeColuna.iColunas + gobjGradeLinha.colGradeCategorias.Count - 1

            'Acumula o total de todas as colunas nessa linhas
            dTotalLinhas = dTotalLinhas + StrParaDbl(GridGrade.TextMatrix(iLinha, iColuna))

            If iLinha = gobjGradeColuna.colGradeCategorias.Count Then

                'Para cada Linha Nessa
                For iLinha2 = gobjGradeColuna.colGradeCategorias.Count To gobjGradeLinha.iLinhas + gobjGradeColuna.colGradeCategorias.Count - 1
                    dTotalColunas = dTotalColunas + StrParaDbl(GridGrade.TextMatrix(iLinha2, iColuna))
                Next

                If dTotalColunas > 0 Then
                    GridGrade.TextMatrix(iLinha2, iColuna) = Formata_Estoque(dTotalColunas)
                Else
                    GridGrade.TextMatrix(iLinha2, iColuna) = ""
                End If
                dTotalColunas = 0
            End If

        Next

        If dTotalLinhas > 0 Then
            GridGrade.TextMatrix(iLinha, iColuna) = Formata_Estoque(dTotalLinhas)
        Else
            GridGrade.TextMatrix(iLinha, iColuna) = ""
        End If

        dTotalGeral = dTotalGeral + dTotalLinhas
        dTotalLinhas = 0
    Next

    GridGrade.TextMatrix(iLinha, iColuna) = Formata_Estoque(dTotalGeral)

    Exit Sub

End Sub

Private Sub BotaoCancela_Click()
    
    'Nao mexer no obj da tela
    giRetornoTela = vbOK
    
    Unload Me
    
    Exit Sub

End Sub

Private Sub BotaoOK_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOK_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 126559
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 126559
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174205)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim objItemRomaneio As ClassItemRomaneioGrade
    
On Error GoTo Erro_Gravar_Registro
    
    'Limpa a coleção de itens vinda da tela chamadora
    Set gobjRomaneioGrade.objObjetoTela.colItensRomaneioGrade = New Collection
    
    'COloca na coleção os dados informados na tela de romaneio
    For Each objItemRomaneio In gobjRomaneioGrade.colItensRomaneioGrade
        'Se há quantidade informada para esse item
        If objItemRomaneio.dQuantidade > 0 Then
        
            'Verifica qual foi a tela que chamou a Grade
            Select Case gobjRomaneioGrade.sNomeTela
            
                Case NOME_TELA_ORDEMPRODUCAO
                    If objItemRomaneio.iAlmoxarifado = 0 Then gError 117627
        
            End Select
        
            'Adiciona ele na colecão
            gobjRomaneioGrade.objObjetoTela.colItensRomaneioGrade.Add objItemRomaneio
        ElseIf objItemRomaneio.dQuantidade = -1 Then
            objItemRomaneio.dQuantidade = 0
            'Adiciona ele na colecão
            gobjRomaneioGrade.objObjetoTela.colItensRomaneioGrade.Add objItemRomaneio
        End If
    Next

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
            
        Case 117627
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO_GRADE", gErr, objItemRomaneio.colCategoria(1).sItem, objItemRomaneio.colCategoria(2).sItem)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174206)

    End Select

    Exit Function

End Function


Private Sub GridGrade_Scroll()
    Call Grid_Scroll(objGridGrade)
End Sub

Private Sub QuantCancelada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub QuantCancelada_GotFocus()
    dQuantCanceladaAnterior = StrParaDbl(QuantCancelada.Text)
End Sub

Private Sub QuantCancelada_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim bQuantCanceladaIgual As Boolean
Dim dQuantCancelada As Double
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_QuantCancelada_Validate
    
    If Len(QuantCancelada.Text) > 0 Then
    
        'Critica o valor da quantidade
        lErro = Valor_NaoNegativo_Critica(QuantCancelada.Text)
        If lErro <> SUCESSO Then gError 26686

        QuantCancelada.Text = Formata_Estoque(QuantCancelada.Text)

    End If
    
    If dQuantCanceladaAnterior = StrParaDbl(QuantCancelada.Text) Then bQuantCanceladaIgual = True
    
    If Not bQuantCanceladaIgual Then

        Set objItemRomaneio = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual)

        dQuantCancelada = StrParaDbl(QuantCancelada.Text)

        If dQuantCancelada > 0 And objItemRomaneio.dQuantidade < dQuantCancelada Then gError 26687
        If dQuantCancelada > 0 And (objItemRomaneio.dQuantidade - dQuantCancelada < objItemRomaneio.dQuantFaturada) Then gError 26688

        lErro = Reserva_Processa(objItemRomaneio.dQuantidade, dQuantCancelada, objItemRomaneio.dQuantFaturada)
        If lErro <> SUCESSO Then gError 26832

        objItemRomaneio.dQuantCancelada = dQuantCancelada
        
    End If
           
          
    Exit Sub
    
Erro_QuantCancelada_Validate:

    Cancel = True

    Select Case gErr
    
        Case 26686, 26832
        
        Case 26687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA", gErr)
        
        Case 26688
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_SUPERIOR", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174207)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Quantidade_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus(Index As Integer)
    Call Grid_Campo_Recebe_Foco(objGridGrade)
End Sub

Private Sub Quantidade_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGrade)
End Sub

Private Sub Quantidade_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridGrade.objControle = Quantidade(Index)
    lErro = Grid_Campo_Libera_Foco(objGridGrade)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Alocacao_Processa(dQuantidade As Double) As Long
 
Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim sUM As String
Dim objItemRomaneiGrade As ClassItemRomaneioGrade

On Error GoTo Erro_Alocacao_Processa

    Set objItemRomaneiGrade = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual)

    objProduto.sCodigo = objItemRomaneiGrade.sProduto
    
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 46950
    If lErro = 28030 Then gError 46952 'Não encontrou
    
    objItemRomaneiGrade.sUMEstoque = objProduto.sSiglaUMEstoque
    
    'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then
        'recolhe a UM do ItemNF
        sUM = UnidadeMed.Caption
        
        If gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_NFISCAL Or ROMANEIOGRADE_FUNCIONAMENTO_RECEBIMENTO Then
        
            'Tenta Alocar o produto no Almoxarifado padrão
            lErro = AlocaAlmoxarifadoPradrao(dQuantidade, objProduto, sUM)
            If lErro <> SUCESSO Then gError 46951
        
        End If
    
    End If

    Alocacao_Processa = SUCESSO

    Exit Function

Erro_Alocacao_Processa:

    Alocacao_Processa = gErr

    Select Case gErr

        Case 46949, 46950, 46951

        Case 46952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174208)

    End Select

    Exit Function

End Function


Function AlocaAlmoxarifadoPradrao(dQuantidade As Double, objProduto As ClassProduto, sUM As String) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim dQuantAlocar As Double
Dim dFator As Double
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dQuantDisponivel As Double
Dim objItemNF As ClassItemNF
Dim colOutrosProdutos As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItemNFAloc As ClassItemNFAlocacao
Dim sProdutoEnxuto As String
Dim dAcrescimo As Double
Dim iNumCasasDec As Integer
Dim dTotal As Double
Dim objItemPV As New ClassItemPedido
Dim colReservaBD As New colReservaItem
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_AlocaAlmoxarifadoPradrao

    'Faz a conversão da UM da tela para a UM de estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 46954


    '####################################################################
    'Alterado por Wagner 16/11/04
'    'Converte a quantidade para a UM de estoque
'    dQuantAlocar = dQuantidade * dFator
'
'    'Calcula o número de casas decimais do Formato de Estoque
'    iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'    If StrParaDbl(Formata_Estoque(dQuantAlocar)) < dQuantAlocar Then
'        dQuantAlocar = StrParaDbl(Formata_Estoque(dQuantAlocar)) + dAcrescimo
'    End If

    dQuantAlocar = Arredonda_Estoque(dQuantidade * dFator)
    '####################################################################
    
    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO Then gError 46956
    
    If iAlmoxarifado = 0 And gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_NFISCAL Then gError 35822
    
    'Se encontrou
    If iAlmoxarifado > 0 Then

        objAlmoxarifado.iCodigo = iAlmoxarifado
        'Lê o Aloxarifado
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> 25056 And lErro <> SUCESSO Then gError 46957
        If lErro = 25056 Then gError 46960
        
        If gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_RECEBIMENTO Then
                    
            Set objReservaItem = New ClassReservaItem

            objReservaItem.iAlmoxarifado = objAlmoxarifado.iCodigo
            objReservaItem.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
            objReservaItem.dQuantidade = Formata_Estoque(dQuantAlocar)
            
            Set gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).colLocalizacao = New Collection
            gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).colLocalizacao.Add objReservaItem
        
            AlocaAlmoxarifadoPradrao = SUCESSO
            Exit Function
        
        End If
        
        objEstoqueProduto.sProduto = objProduto.sCodigo
        objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
        'Le os estoques desse produto nesse almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 46955
        If lErro = 21306 Then gError 46961 'Não encontrou
        
        If gobjRomaneioGrade.iModoFuncionamento = ROMANEIOGRADE_FUNCIONAMENTO_NFISCAL Then

            'Seleciona a origem da quantidade disponível
            Select Case gobjRomaneioGrade.iTipoNFiscal
            
                'Se o tipo da nota for cobrança de mat. consignado
                Case DOCINFO_NFISPC, DOCINFO_NFFISPC
                    
                    'A quantidade disponível deve ser igual a quantidade no escaninho de mat. em Consignação (Consig)
                    dQuantDisponivel = objEstoqueProduto.dQuantConsig
                
                'Se o tipo da nota for mat. beneficiado de 3º´s
                Case DOCINFO_NFISBF, DOCINFO_NFISFBF
                
                    'A quantidade disponível deve ser igual a quantidade no escaninho Mat.de 3º´s em Beneficiamento (Benef3)
                    dQuantDisponivel = objEstoqueProduto.dQuantBenef3
                    
                'Se for outro tipo de nota
                Case Else
                    
                    'A quantidade disponível deve ser igual a quantidade do escaninho mat. nosso disponível (DispNossa)
                    dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
            
            End Select

        Else

            objItemPV.lNumIntDoc = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).lNumIntItemPV
            objItemPV.sProduto = objProduto.sCodigo
            objItemPV.iPossuiGrade = MARCADO
    
            lErro = CF("ReservasItemPV_Le_NumIntOrigem", objItemPV, colReservaBD)
            If lErro <> SUCESSO And lErro <> 51601 Then gError 62095
            
            For iIndice = 1 To colReservaBD.Count
                If objEstoqueProduto.iAlmoxarifado = colReservaBD(iIndice).iAlmoxarifado Then
                    objEstoqueProduto.dQuantDispNossa = objEstoqueProduto.dQuantDispNossa + colReservaBD(iIndice).dQuantidade
                    Exit For
                End If
            Next
            
            dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
        
        End If
        

        'Verifica se a Quantidade disponível é maior que a quantidade a alocar
        If (dQuantAlocar - dQuantDisponivel) < QTDE_ESTOQUE_DELTA Then

            Set objReservaItem = New ClassReservaItem

            objReservaItem.iAlmoxarifado = objAlmoxarifado.iCodigo
            objReservaItem.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
            objReservaItem.dQuantidade = Formata_Estoque(dQuantAlocar)
            
            Set gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).colLocalizacao = New Collection
            gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).colLocalizacao.Add objReservaItem
            
        'Se não for
        Else
                    
            Set objItemNF = New ClassItemNF
            'Recolhe os dados do item
            objItemNF.iItem = gobjRomaneioGrade.objObjetoTela.iItem
            objItemNF.sProduto = objProduto.sCodigo
            objItemNF.sDescricaoItem = objProduto.sDescricao
            objItemNF.dQuantidade = dQuantidade
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.lNumIntItemPedVenda = gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).lNumIntItemPV
            objItemNF.iPossuiGrade = MARCADO

            'Recolhe todos os produtos dos outros itens
            For iIndice = 1 To gobjRomaneioGrade.colItensRomaneioGrade.Count
                If iIndice <> gobjRomaneioGrade.iItemAtual Then
                    'Adiciona na coleção de produtos
                    colOutrosProdutos.Add gobjRomaneioGrade.colItensRomaneioGrade(iIndice).sProduto
                End If
            Next

            'Chama a tela de Localização de Produto
            Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, DOCINFO_NFISFVPV)
            If giRetornoTela = vbCancel Then gError 46963 'Se nada foi feito lá
            If giRetornoTela = vbOK Then

                'Se o produto foi substituido
                If objProduto.sCodigo <> objItemNF.sProduto Then gError 46962
                                
                Set gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).colLocalizacao = New Collection
                
                'Calcula o Total Alocado
                For Each objItemNFAloc In objItemNF.colAlocacoes
                    dTotal = dTotal + objItemNFAloc.dQuantidade
                Next
                
                'Para cada alocação feita para o item
                For Each objItemNFAloc In objItemNF.colAlocacoes

                    Set objReservaItem = New ClassReservaItem
        
                    objReservaItem.iAlmoxarifado = objItemNFAloc.iAlmoxarifado
                    objReservaItem.sAlmoxarifado = objItemNFAloc.sAlmoxarifado
                    objReservaItem.dQuantidade = objItemNFAloc.dQuantidade
                

                    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).colLocalizacao.Add objReservaItem
                Next
            
                Quantidade(GridGrade.Col) = Formata_Estoque(dTotal)

            End If
        End If
        
    End If

    AlocaAlmoxarifadoPradrao = SUCESSO

    Exit Function

Erro_AlocaAlmoxarifadoPradrao:

    AlocaAlmoxarifadoPradrao = gErr

    Select Case gErr

        Case 46954, 46955, 46956, 46957, 46958, 46959, 62095, 35822

        Case 46960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 46961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 46962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO", gErr)

        Case 46963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174209)

        End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoLocalizacao = Nothing
    Set objEventoVersao = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoEstoqueProdSai = Nothing
    Set objEventoOP = Nothing
    Set objEventoProdutoOP = Nothing
    Set objEventoOPProd = Nothing
    Set objEventoEstoqueProd = Nothing
    
    Set gobjGradeLinha = Nothing
    Set gobjGradeColuna = Nothing
    
    Set gobjRomaneioGrade = Nothing
    Set objGridGrade = Nothing

End Sub

Private Sub Versao_Click()

    iAlterado = REGISTRO_ALTERADO

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sVersao = Versao.Text

End Sub

Private Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 78739

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOPProdSai.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOPProdSai.ItemData(FilialOPProdSai.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 78739 'Erro já tratado
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174211)

    End Select

    Exit Function

End Function

Private Sub CodOPProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodOPProd_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemOP As New ClassItemOP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iCont As Integer
Dim objItemOPUnico As ClassItemOP
Dim sProdutoOPEnxuto As String
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_CodOpProd_Validate

    If Len(Trim(CodOPProd.Text)) > 0 Then

        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = CodOPProd.Text

        lErro = CF("OrdemProducao_Le", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 126680

        If lErro = 30368 Then gError 126681

        'ordem de producao baixada
        If lErro = 55316 Then gError 126682

        'preenche o almoxarifado no grid a partir do item da OP
        lErro = Preenche_Almoxarifado_Prod(giFilialEmpresa, CodOPProd.Text, gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto, objAlmoxarifado)
        If lErro <> SUCESSO Then gError 126683

        AlmoxProd.Text = objAlmoxarifado.sNomeReduzido

        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = objAlmoxarifado.iCodigo
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = objAlmoxarifado.sNomeReduzido

    Else

        AlmoxProd.Text = ""
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = 0
        gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = ""
        
    End If

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sCodOP = CodOPProd.Text
          
    Exit Sub

Erro_CodOpProd_Validate:

    Cancel = True

    Select Case gErr

        Case 126680, 126683

        Case 126681
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_OPCODIGO_NAO_CADASTRADO", objOrdemProducao.sCodigo)

            If vbMsg = vbYes Then

                Call Chama_Tela_Modal("OrdemProducao", objOrdemProducao)

            End If

        Case 126682
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174212)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Almoxarifado_Prod(ByVal iFilialEmpresa As Integer, ByVal sOPCodigo As String, ByVal sProduto As String, objAlmoxarifado As ClassAlmoxarifado) As Long
'preenche o almoxarifado no grid a partir do item da OP

Dim objItemOP As New ClassItemOP
Dim lErro As Long

On Error GoTo Erro_Preenche_Almoxarifado_Prod

    lErro = CF("Cust_RomaneioGrade_PreencheAlmox", sProduto, gobjRomaneioGrade)
    If lErro <> SUCESSO Then gError 133019

    objItemOP.iFilialEmpresa = iFilialEmpresa
    objItemOP.sCodigo = sOPCodigo
    objItemOP.sProduto = sProduto

    lErro = CF("ItemOP_Le", objItemOP)
    If lErro <> SUCESSO And lErro <> 34711 Then gError 126684

    If lErro = 34711 Then gError 126685
    
    objAlmoxarifado.iCodigo = objItemOP.iAlmoxarifado
    
    'le o nome reduzido do almoxarifado associado ao itemop
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 126686
    
    If lErro = 25056 Then gError 126687
    
    Preenche_Almoxarifado_Prod = SUCESSO

    Exit Function

Erro_Preenche_Almoxarifado_Prod:

    Preenche_Almoxarifado_Prod = gErr
    
    Select Case gErr
    
        Case 126684, 126686
    
        Case 126685
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 126687
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objItemOP.iAlmoxarifado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174213)
    
    End Select
    
    Exit Function

End Function

Private Sub LabelOPProd_Click()

Dim objOrdemProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection

On Error GoTo Erro_LabelOPProd_Click
    
    Call Chama_Tela_Modal("OrdemProducaoLista", colSelecao, objOrdemProducao, objEventoOPProd)
   
    Exit Sub

Erro_LabelOPProd_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174214)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOPProd_evSelecao(obj1 As Object)

Dim objOrdemProducao As ClassOrdemDeProducao
    
On Error GoTo Erro_objEventoOPProd_evSelecao

    Set objOrdemProducao = obj1
    
    CodOPProd.Text = objOrdemProducao.sCodigo
    
    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sCodOP = objOrdemProducao.sCodigo
        
    Me.Show

    Exit Sub
    
Erro_objEventoOPProd_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174215)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub AlmoxProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AlmoxProd_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AlmoxProd_Validate

    lErro = AlmoxGeral_Validate(AlmoxProd)
    If lErro <> SUCESSO Then gError 126688

    Exit Sub

Erro_AlmoxProd_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 126688

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174216)

    End Select

    Exit Sub

End Sub

Private Sub LabelAlmoxProd_Click()
'Informa se produto é estocado em algum almoxarifado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelAlmoxProd_Click

    colSelecao.Add gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sProduto
    
    'chama a tela de lista de estoque do produto corrente
    Call Chama_Tela_Modal("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoqueProd)
    
    Exit Sub

Erro_LabelAlmoxProd_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174217)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoqueProd_evselecao(obj1 As Object)

Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lErro As Long

On Error GoTo Erro_objEventoEstoqueProd_evselecao

    Set objEstoqueProduto = obj1

    AlmoxProd.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).iAlmoxarifado = objEstoqueProduto.iAlmoxarifado
    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sAlmoxarifado = objEstoqueProduto.sAlmoxarifadoNomeReduzido
    
    Me.Show

    Exit Sub

Erro_objEventoEstoqueProd_evselecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174218)

    End Select

    Exit Sub

End Sub

Private Sub HorasMaqProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub HorasMaqProd_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HorasMaqProd_Validate

    If Len(Trim(HorasMaqProd.Text)) <> 0 Then

        lErro = Valor_Positivo_Critica(HorasMaqProd.Text)
        If lErro <> SUCESSO Then gError 126689

    End If

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).lHorasMaquina = StrParaLong(HorasMaqProd.Text)

    Exit Sub

Erro_HorasMaqProd_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 126689
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 174219)

    End Select

    Exit Sub

End Sub

Private Sub LoteProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LoteProd_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim iEntradaSaida As Integer

On Error GoTo Erro_LoteProd_Validate

    gobjRomaneioGrade.colItensRomaneioGrade(gobjRomaneioGrade.iItemAtual).sLote = LoteProd.Text

    Exit Sub

Erro_LoteProd_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174220)

    End Select

    Exit Sub

End Sub

Public Function Grid_Monta_Colunas(ByVal objGradeColuna As ClassGradeLinCol, ByVal objGrid As AdmGrid, ByVal iPosInicial As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objGradeLinColCat As ClassGradeLinColCat
Dim objGradeLinColCatItem As ClassGradeLinColCatItem
Dim objGradeLinColCatAux As ClassGradeLinColCat
Dim objGradeLinColCatItemAux As ClassGradeLinColCatItem
Dim colColunas As Collection
Dim iNumNiveis As Integer
Dim iNumColunas As Integer
Dim iNumRepeticoes As Integer

On Error GoTo Erro_Grid_Monta_Colunas

    iNumNiveis = objGradeColuna.colGradeCategorias.Count
    
    If iPosInicial = 0 Then iPosInicial = 1

    For iIndice = iNumNiveis To 2 Step -1
    
        Set objGradeLinColCat = objGradeColuna.colGradeCategorias.Item(iIndice)
        Set objGradeLinColCatAux = objGradeColuna.colGradeCategorias.Item(iIndice - 1)
        
        If iIndice = iNumNiveis Then
            objGradeLinColCat.iNumRepeticoes = 1
        End If
        
        For Each objGradeLinColCatItemAux In objGradeLinColCatAux.colItens
                
            Set colColunas = New Collection
                       
            iNumRepeticoes = 0
            For Each objGradeLinColCatItem In objGradeLinColCat.colItens
                iNumRepeticoes = iNumRepeticoes + 1
                colColunas.Add objGradeLinColCatItem
                iNumColunas = iNumColunas + 1
            Next
            
            objGradeLinColCatAux.iNumRepeticoes = objGradeLinColCat.iNumRepeticoes * iNumRepeticoes
                        
            Set objGradeLinColCatItemAux.colItens = colColunas
        
        Next

    Next
    
    If iNumNiveis = 1 Then
        iNumColunas = objGradeColuna.colGradeCategorias.Item(1).colItens.Count
        objGradeColuna.colGradeCategorias.Item(1).iNumRepeticoes = 1
    End If
    
    objGradeColuna.iColunas = iNumColunas
    
    For iIndice = 1 To iNumColunas + iPosInicial
        objGrid.colColuna.Add (" ")
    Next
    
    GridGrade.FixedRows = objGradeColuna.colGradeCategorias.Count
       
    Grid_Monta_Colunas = SUCESSO

    Exit Function

Erro_Grid_Monta_Colunas:

    Grid_Monta_Colunas = gErr

    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174206)

    End Select

    Exit Function

End Function

Public Function Grid_Exibe_Colunas(ByVal objGradeColuna As ClassGradeLinCol, ByVal iPosInicial As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objGradeLinColCat As ClassGradeLinColCat
Dim objGradeLinColCatItem As ClassGradeLinColCatItem
Dim iLinha As Integer
Dim iPos As Integer

On Error GoTo Erro_Grid_Exibe_Colunas
      
    iLinha = -1
    For Each objGradeLinColCat In objGradeColuna.colGradeCategorias
    
        iLinha = iLinha + 1
        
        iPos = iPosInicial - 1
            
        Do While (iPos - (iPosInicial - 1)) <> objGradeColuna.iColunas
        
            For Each objGradeLinColCatItem In objGradeLinColCat.colItens
            
                For iIndice = 1 To objGradeLinColCat.iNumRepeticoes
                
                    iPos = iPos + 1
                    GridGrade.TextMatrix(iLinha, iPos) = objGradeLinColCatItem.objCategoriaProdutoItem.sItem
                
                Next
            
            Next
            
            If (iPos - (iPosInicial - 1)) > objGradeColuna.iColunas Then gError 180101
            
        Loop
    
    Next
    
    Grid_Exibe_Colunas = SUCESSO

    Exit Function

Erro_Grid_Exibe_Colunas:

    Grid_Exibe_Colunas = gErr

    Select Case gErr
    
        Case 180101
            Call Rotina_Erro(vbOKOnly, "ERRO_EXIBICAO_COLUNAS_GRADE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180102)

    End Select

    Exit Function

End Function

Public Function Grid_Monta_Linhas(ByVal objGradeLinha As ClassGradeLinCol, ByVal objGrid As AdmGrid, ByVal iNumColunas As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objGradeLinColCat As ClassGradeLinColCat
Dim objGradeLinColCatItem As ClassGradeLinColCatItem
Dim objGradeLinColCatAux As ClassGradeLinColCat
Dim objGradeLinColCatItemAux As ClassGradeLinColCatItem
Dim colLinhas As Collection
Dim iNumNiveis As Integer
Dim iNumLinhas As Integer
Dim iNumRepeticoes As Integer

On Error GoTo Erro_Grid_Monta_Linhas

    iNumNiveis = objGradeLinha.colGradeCategorias.Count

    For iIndice = iNumNiveis To 2 Step -1
    
        Set objGradeLinColCat = objGradeLinha.colGradeCategorias.Item(iIndice)
        Set objGradeLinColCatAux = objGradeLinha.colGradeCategorias.Item(iIndice - 1)
        
        If iIndice = iNumNiveis Then
            objGradeLinColCat.iNumRepeticoes = 1
        End If
        
        For Each objGradeLinColCatItemAux In objGradeLinColCatAux.colItens
                
            Set colLinhas = New Collection
                       
            iNumRepeticoes = 0
            For Each objGradeLinColCatItem In objGradeLinColCat.colItens
                iNumRepeticoes = iNumRepeticoes + 1
                colLinhas.Add objGradeLinColCatItem
                iNumLinhas = iNumLinhas + 1
            Next
            
            objGradeLinColCatAux.iNumRepeticoes = objGradeLinColCat.iNumRepeticoes * iNumRepeticoes
                        
            Set objGradeLinColCatItemAux.colItens = colLinhas
        
        Next

    Next
    
    If iNumNiveis = 1 Then
        iNumLinhas = objGradeLinha.colGradeCategorias.Item(1).colItens.Count
        objGradeLinha.colGradeCategorias.Item(1).iNumRepeticoes = 1
    End If
    
    If iNumNiveis = 0 Then
        Set objGradeLinColCat = New ClassGradeLinColCat
        Set objGradeLinColCatItem = New ClassGradeLinColCatItem
        objGradeLinColCat.iNumRepeticoes = 1
        objGradeLinColCat.colItens.Add objGradeLinColCatItem
        objGradeLinha.colGradeCategorias.Add objGradeLinColCat
        iNumLinhas = 1
        iNumNiveis = 1
    End If
    
    objGradeLinha.iLinhas = iNumLinhas
    
    For iIndice = 1 To iNumColunas + iNumNiveis
    
        'Inclui na tela um novo Controle para essa coluna
        Load Quantidade(iIndice)
        
        'Traz o controle recem desenhado para a frente
        Quantidade(iIndice).ZOrder
        'Torna o controle visível
        Quantidade(iIndice).Visible = True
        'Informa ao objGrid que esse controle é dessa coluna
        objGrid.colCampo.Add (Quantidade(iIndice).Name)
        'Informa ao objGrid qual é o indice desse controle no array criado
        objGrid.colIndex.Add iIndice
        
    Next
    
    GridGrade.FixedCols = objGradeLinha.colGradeCategorias.Count
       
    Grid_Monta_Linhas = SUCESSO

    Exit Function

Erro_Grid_Monta_Linhas:

    Grid_Monta_Linhas = gErr

    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174206)

    End Select

    Exit Function

End Function

Public Function Grid_Exibe_Linhas(ByVal objGradeLinha As ClassGradeLinCol, ByVal iPosInicial As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objGradeLinColCat As ClassGradeLinColCat
Dim objGradeLinColCatItem As ClassGradeLinColCatItem
Dim iColuna As Integer
Dim iPos As Integer

On Error GoTo Erro_Grid_Exibe_Linhas
      
    iColuna = -1
    For Each objGradeLinColCat In objGradeLinha.colGradeCategorias
    
        iColuna = iColuna + 1
        
        iPos = iPosInicial - 1
            
        For iIndice = 1 To objGridGrade.iLinhasExistentes + GridGrade.FixedRows - 1
            GridGrade.TextMatrix(iIndice, iColuna) = " "
        Next
            
        Do While (iPos - (iPosInicial - 1)) <> objGradeLinha.iLinhas
        
            For Each objGradeLinColCatItem In objGradeLinColCat.colItens
            
                For iIndice = 1 To objGradeLinColCat.iNumRepeticoes
                
                    iPos = iPos + 1
                    GridGrade.TextMatrix(iPos, iColuna) = objGradeLinColCatItem.objCategoriaProdutoItem.sItem
                    GridGrade.ColAlignment(iColuna) = vbAlignRight
                
                Next
            
            Next
            
            If (iPos - (iPosInicial - 1)) > objGradeLinha.iLinhas Then gError 180103
            
        Loop
    
    Next
    
    Grid_Exibe_Linhas = SUCESSO

    Exit Function

Erro_Grid_Exibe_Linhas:

    Grid_Exibe_Linhas = gErr

    Select Case gErr
            
        Case 180103
            Call Rotina_Erro(vbOKOnly, "ERRO_EXIBICAO_LINHAS_GRADE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180104)

    End Select

    Exit Function

End Function

Public Function Verifica_Romaneio_Grid(ByVal objItemRomaneio As ClassItemRomaneioGrade, ByVal iLinha As Integer, ByVal iColuna As Integer) As Boolean

Dim lErro As Long
Dim objGradeLinColCat As ClassGradeLinColCat
Dim bLinhaOK As Boolean
Dim bColunaOK As Boolean
Dim iLinhaC As Integer
Dim iColunaC As Integer
Dim bOK As Boolean
Dim sItem As String
Dim objCategProdItem As ClassCategoriaProdutoItem
Dim bAchou As Boolean

    bOK = False

    If Not (gobjGradeColuna Is Nothing) And Not (gobjGradeLinha Is Nothing) Then
    
        bLinhaOK = False
        bColunaOK = False

        iLinhaC = 0
        For Each objGradeLinColCat In gobjGradeColuna.colGradeCategorias
    
            sItem = GridGrade.TextMatrix(iLinhaC, iColuna)
            
            bAchou = False
            For Each objCategProdItem In objItemRomaneio.colCategoria
                If UCase(objCategProdItem.sCategoria) = UCase(objGradeLinColCat.objGradeCategoria.sCategoria) And UCase(objCategProdItem.sItem) = UCase(sItem) Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
                bColunaOK = False
                Exit For
            Else
                bColunaOK = True
            End If
        
            iLinhaC = iLinhaC + 1
        Next
          
        iColunaC = 0
        For Each objGradeLinColCat In gobjGradeLinha.colGradeCategorias
        
            sItem = GridGrade.TextMatrix(iLinha, iColunaC)
            
            bAchou = False
            For Each objCategProdItem In objItemRomaneio.colCategoria
                If (UCase(objCategProdItem.sCategoria) = UCase(objGradeLinColCat.objGradeCategoria.sCategoria) And UCase(objCategProdItem.sItem) = UCase(sItem)) Or (objGradeLinColCat.objGradeCategoria.sCategoria = "" And sItem = "") Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
                bLinhaOK = False
                Exit For
            Else
                bLinhaOK = True
            End If
        
            iColunaC = iColunaC + 1
        Next
        
        If bLinhaOK And bColunaOK Then
            bOK = True
        Else
            bOK = False
        End If
    
    End If
    
    Verifica_Romaneio_Grid = bOK
    
End Function

Public Sub Trata_Tecla_Grid(iKey As Integer)

Dim objGradeLinColCat As ClassGradeLinColCat
Dim objGradeLinColCatItem As ClassGradeLinColCatItem
Dim iPosicao As Integer
Dim bAchou As Boolean
Dim iPosicaoAtual As Integer
Dim iIndice As Integer
Dim iPosicaoAtualAux As Integer
Dim iCont As Integer
    
    If Not (gobjGradeColuna Is Nothing) And Not (gobjGradeLinha Is Nothing) Then
    
        iPosicao = gobjGradeLinha.colGradeCategorias.Count
        iPosicaoAtualAux = objGridGrade.objGrid.Col
        iPosicaoAtual = objGridGrade.objGrid.Col
        
        bAchou = False
        iCont = 0
        
        Do While Not bAchou And iCont < 2
        
            iCont = iCont + 1
        
            For Each objGradeLinColCat In gobjGradeColuna.colGradeCategorias
            
                For Each objGradeLinColCatItem In objGradeLinColCat.colItens
                    If Len(objGradeLinColCatItem.objCategoriaProdutoItem.sItem) > 0 Then
                        If Asc(left(objGradeLinColCatItem.objCategoriaProdutoItem.sItem, 1)) = iKey Then
                        
                            If iPosicao > iPosicaoAtualAux Then
                                bAchou = True
                                Exit For
                            End If
                        End If
                    End If
                    iPosicao = iPosicao + objGradeLinColCat.iNumRepeticoes
                Next
        
                If bAchou Then Exit For
        
            Next
                    
            If bAchou Then
                objGridGrade.objGrid.Col = iPosicao
                Call GridGrade_GotFocus
                Call GridGrade_Click
                Call GridGrade_KeyPress(vbKeyReturn)
                
            Else
                If iPosicaoAtualAux > gobjGradeLinha.colGradeCategorias.Count Then
                    iPosicaoAtualAux = gobjGradeLinha.colGradeCategorias.Count - 1
                    iPosicao = gobjGradeLinha.colGradeCategorias.Count
                End If
            End If
            
        Loop
        
    End If
    
End Sub

Public Sub VisualizarFigura(ByVal sProduto As String)

Dim lErro As Long
Dim sFigura As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_VisualizarFigura

    If Len(Trim(sProduto)) > 0 Then
    
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
        sFigura = objProduto.sFigura
        
    End If

    'verifica se a figura foi preenchida
    If Len(Trim(sFigura)) > 0 Then
        Figura.Picture = LoadPicture(sFigura)
    Else
        Figura.Picture = LoadPicture
    End If
    
    
       
    Exit Sub
    
Erro_VisualizarFigura:

    Exit Sub

End Sub
