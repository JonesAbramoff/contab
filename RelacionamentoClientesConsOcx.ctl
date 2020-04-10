VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelacionamentoClientesConsOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   2
      Left            =   195
      TabIndex        =   24
      Top             =   750
      Visible         =   0   'False
      Width           =   9030
      Begin VB.Frame FrameRelacionamentos 
         Caption         =   "Relacionamentos"
         Height          =   5040
         Left            =   0
         TabIndex        =   25
         Top             =   15
         Width           =   9045
         Begin VB.TextBox FilialClienteGrid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6960
            TabIndex        =   28
            Top             =   960
            Width           =   825
         End
         Begin VB.TextBox Status 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4320
            TabIndex        =   38
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox ClienteGrid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5760
            TabIndex        =   37
            Top             =   960
            Width           =   1530
         End
         Begin VB.TextBox Fone 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3000
            TabIndex        =   27
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Atendente 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2160
            TabIndex        =   26
            Top             =   960
            Width           =   1245
         End
         Begin MSMask.MaskEdBox OrigemGrid 
            Height          =   225
            Left            =   990
            TabIndex        =   29
            Top             =   960
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contato 
            Height          =   225
            Left            =   1560
            TabIndex        =   30
            Top             =   1560
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipoRelacionamento 
            Height          =   225
            Left            =   3960
            TabIndex        =   31
            Top             =   960
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   225
            Left            =   3240
            TabIndex        =   32
            Top             =   960
            Visible         =   0   'False
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   225
            Left            =   240
            TabIndex        =   33
            Top             =   1395
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridRelacionamentos 
            Height          =   1815
            Left            =   45
            TabIndex        =   34
            Top             =   210
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   3201
            _Version        =   393216
         End
         Begin VB.Label LabelAssunto 
            AutoSize        =   -1  'True
            Caption         =   "Assunto:"
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
            Left            =   60
            TabIndex        =   36
            Top             =   3645
            Width           =   750
         End
         Begin VB.Label Assunto 
            BorderStyle     =   1  'Fixed Single
            Height          =   1095
            Left            =   60
            TabIndex        =   35
            Top             =   3885
            Width           =   8910
         End
      End
   End
   Begin VB.CheckBox OpcaoPadrao 
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
      Left            =   4320
      TabIndex        =   58
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   5025
      Index           =   1
      Left            =   250
      TabIndex        =   23
      Top             =   765
      Width           =   8970
      Begin VB.Frame FrameSelecao 
         Caption         =   "Seleção"
         Height          =   3735
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   8115
         Begin VB.Frame FrameCliente 
            Caption         =   "Cliente"
            Height          =   1380
            Left            =   2280
            TabIndex        =   50
            Top             =   360
            Width           =   3315
            Begin VB.ComboBox FilialCliente 
               Height          =   315
               Left            =   915
               TabIndex        =   10
               Top             =   840
               Width           =   2145
            End
            Begin MSMask.MaskEdBox Cliente 
               Height          =   300
               Left            =   900
               TabIndex        =   8
               Top             =   360
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label LabelCliente 
               AutoSize        =   -1  'True
               Caption         =   "Cliente:"
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
               Left            =   195
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   7
               Top             =   405
               Width           =   660
            End
            Begin VB.Label LabelFilialCliente 
               AutoSize        =   -1  'True
               Caption         =   "Filial:"
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
               Left            =   390
               TabIndex        =   9
               Top             =   885
               Width           =   465
            End
         End
         Begin VB.Frame FrameStatus 
            Caption         =   "Status"
            Height          =   1380
            Left            =   5640
            TabIndex        =   49
            Top             =   360
            Width           =   2355
            Begin VB.OptionButton StatusTodos 
               Caption         =   "Todos"
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
               TabIndex        =   13
               Top             =   960
               Width           =   855
            End
            Begin VB.OptionButton StatusEncerrado 
               Caption         =   "Encerrado"
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
               TabIndex        =   12
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton StatusPendente 
               Caption         =   "Pendente"
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
               TabIndex        =   11
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Frame FrameCodigo 
            Caption         =   "Código"
            Height          =   825
            Left            =   -20040
            TabIndex        =   44
            Top             =   240
            Width           =   3795
            Begin MSMask.MaskEdBox CodigoDe 
               Height          =   300
               Left            =   600
               TabIndex        =   45
               ToolTipText     =   "Informe o código do relacionamento."
               Top             =   360
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               Mask            =   "999999999999"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoAte 
               Height          =   300
               Left            =   2400
               TabIndex        =   46
               ToolTipText     =   "Informe o código do relacionamento."
               Top             =   360
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               Mask            =   "999999999999"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoAte 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               Left            =   1995
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   48
               Top             =   390
               Width           =   360
            End
            Begin VB.Label LabelCodigoDe 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               Height          =   240
               Left            =   195
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   47
               Top             =   390
               Width           =   345
            End
         End
         Begin VB.Frame FrameOrigem 
            Caption         =   "Origem"
            Height          =   1380
            Left            =   5640
            TabIndex        =   43
            Top             =   1920
            Width           =   2355
            Begin VB.ComboBox Origem 
               Height          =   315
               ItemData        =   "RelacionamentoClientesConsOcx.ctx":0000
               Left            =   840
               List            =   "RelacionamentoClientesConsOcx.ctx":000D
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label LabelOrigem 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   600
               Width           =   660
            End
         End
         Begin VB.Frame FramePeriodo 
            Caption         =   "Período"
            Height          =   1380
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   2115
            Begin MSComCtl2.UpDown UpDownPeriodoAte 
               Height          =   300
               Left            =   1680
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   840
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox PeriodoAte 
               Height          =   300
               Left            =   720
               TabIndex        =   5
               Top             =   840
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownPeriodoDe 
               Height          =   300
               Left            =   1680
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox PeriodoDe 
               Height          =   300
               Left            =   720
               TabIndex        =   2
               Top             =   360
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label LabelDataAte 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               Left            =   240
               TabIndex        =   4
               Top             =   900
               Width           =   360
            End
            Begin VB.Label LabelDataDe 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               TabIndex        =   1
               Top             =   420
               Width           =   315
            End
         End
         Begin VB.Frame FrameTipo 
            Caption         =   "Tipo"
            Height          =   1380
            Left            =   2280
            TabIndex        =   41
            Top             =   1920
            Width           =   3315
            Begin VB.ComboBox Tipo 
               Height          =   315
               Left            =   1200
               TabIndex        =   20
               Top             =   915
               Width           =   1995
            End
            Begin VB.OptionButton TipoApenas 
               Caption         =   "Apenas"
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
               TabIndex        =   19
               Top             =   945
               Width           =   1050
            End
            Begin VB.OptionButton TipoTodos 
               Caption         =   "Todos os tipos"
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
               Left            =   195
               TabIndex        =   18
               Top             =   360
               Value           =   -1  'True
               Width           =   1620
            End
         End
         Begin VB.Frame FrameAtendentes 
            Caption         =   "Atendentes"
            Height          =   1380
            Left            =   120
            TabIndex        =   40
            Top             =   1920
            Width           =   2115
            Begin VB.ComboBox AtendenteAte 
               Height          =   315
               Left            =   600
               TabIndex        =   17
               Top             =   945
               Width           =   1335
            End
            Begin VB.ComboBox AtendenteDe 
               Height          =   315
               Left            =   600
               TabIndex        =   15
               Top             =   345
               Width           =   1335
            End
            Begin VB.Label LabelAtendenteAte 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               Left            =   195
               TabIndex        =   16
               Top             =   1005
               Width           =   360
            End
            Begin VB.Label LabelAtendenteDe 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               Left            =   195
               TabIndex        =   14
               Top             =   405
               Width           =   315
            End
         End
      End
   End
   Begin VB.ComboBox OpcoesTela 
      Height          =   315
      Left            =   1320
      TabIndex        =   56
      Top             =   60
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6780
      ScaleHeight     =   495
      ScaleWidth      =   2565
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   30
      Width           =   2625
      Begin VB.CommandButton BotaoAtualizar 
         Height          =   360
         Left            =   120
         Picture         =   "RelacionamentoClientesConsOcx.ctx":002A
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Atualiza a lista de relacionamentos"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1080
         Picture         =   "RelacionamentoClientesConsOcx.ctx":07EC
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   600
         Picture         =   "RelacionamentoClientesConsOcx.ctx":0976
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1560
         Picture         =   "RelacionamentoClientesConsOcx.ctx":0AD0
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2040
         Picture         =   "RelacionamentoClientesConsOcx.ctx":1002
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5520
      Left            =   60
      TabIndex        =   0
      Top             =   375
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   9737
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Relacionamentos"
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
   Begin VB.Label LabelOpcao 
      AutoSize        =   -1  'True
      Caption         =   "Opção:"
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
      Left            =   600
      TabIndex        =   57
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "RelacionamentoClientesConsOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos de browser
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoAtendente As AdmEvento
Attribute objEventoAtendente.VB_VarHelpID = -1

'Coleção usada para manter os relacionamentos do grid na memória
'Isso é feito apenas para facilitar a exibição do campo assunto
'E obtenção dos códigos de atendentes, clientes, filiais, tipos, etc...
Dim gcolRelacionamentos As Collection

Dim iAlterado As Integer
Dim iAtualizaGrid As Integer
Dim iClienteAlterado As Integer

'Usada para evitar que a função OpcaoClick seja disparada erroneamente
Public iAtualizaTela As Integer

'Controle do frame exibido
Dim giFrameAtual As Integer

'Variáveis para controle do evento seleção
Dim giEventoCodigo As Integer

'Constantes para controle do evento seleção
Const EVENTO_CODIGODE = 1
Const EVENTO_CODIGOATE = 2

'Variáveis usadas no controle do grid
Dim objGridRelacionamentos As New AdmGrid

Dim iGrid_Data_Col As Integer
Dim iGrid_Origem_Col As Integer
Dim iGrid_Atendente_Col As Integer
Dim iGrid_TipoRelacionamento_Col As Integer
Dim iGrid_Codigo_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_FilialCliente_Col As Integer
Dim iGrid_Contato_Col As Integer
Dim iGrid_Telefone_Col As Integer
Dim iGrid_Status_Col As Integer

Const NUM_MAX_RELACIONAMENTOS = 1000

'Constantes para identificação de tabs
Const TAB_Selecao = 1
Const TAB_RELACIONAMENTOS = 2

'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Consulta de Relacionamento com clientes"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "RelacionamentoClientesCons"
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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub
'***************************************************
'Fim Trecho de codigo comum as telas
'***************************************************

'*** CARREGAMENTO DA TELA - INÍCIO ***
Private Function Form_Load()

Dim lErro As Long
Dim objTela As Object

On Error GoTo Erro_Form_Load

    giFrameAtual = 1
    iAtualizaTela = 1 'indica que a função opcoestela_click deve ser executada
    
    'Inicializa eventos de browser
    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoAtendente = New AdmEvento

    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo, False)
    If lErro <> SUCESSO Then gError 102792
    
    'Carrega a combo AtendenteDe
    lErro = CF("Carrega_Atendentes", AtendenteDe)
    If lErro <> SUCESSO Then gError 102793
    
    'Preenche valores default da tela
    'Data Inicial
    PeriodoDe.PromptInclude = False
    PeriodoDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    PeriodoDe.PromptInclude = True
    
    'Data Final
    PeriodoAte.PromptInclude = False
    PeriodoAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    PeriodoAte.PromptInclude = True
    
    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", AtendenteAte)
    If lErro <> SUCESSO Then gError 102794
    
    'Origem
    Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
    
    'Executa inicializacao do GridRelacionamentos
    lErro = Inicializa_GridRelacionamentos(objGridRelacionamentos)
    If lErro <> SUCESSO Then gError 102814

    'Guarda em objTela os dados dessa tela
    Set objTela = Me
    
    '*** CARREGAMENTO DA COMBO OpcoesTela - INÍCIO ***
    'Esse carregamento é feito aqui para garantir que, caso seja selecionada uma opção padrão,
    'os valores dessa opção sejam exibidos normalmente na tela
    'Carrega a combo de opçõs
    lErro = CF("Carrega_OpcoesTela", objTela, True)
    If lErro <> SUCESSO Then gError 102894
    '*** CARREGAMENTO DA COMBO OpcoesTela - FIM ***
    
    iAlterado = 0
    iClienteAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Function

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 102792 To 102794, 102894

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166626)

    End Select

    iAlterado = 0
    iClienteAlterado = 0
    
End Function

Public Function Trata_Parametros(Optional ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166627)

    End Select

End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Private Sub Unload(objme As Object)
    RaiseEvent Unload
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    Set objEventoAtendente = Nothing
    Set gcolRelacionamentos = Nothing

    Call ComandoSeta_Liberar(Me.Name)
    
End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub CodigoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodigoDe, iAlterado)
End Sub

Private Sub CodigoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodigoAte, iAlterado)
End Sub

Private Sub PeriodoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(PeriodoDe, iAlterado)
End Sub

Private Sub PeriodoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(PeriodoAte, iAlterado)
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub OpcoesTela_Click()
    
Dim lErro As Long
Dim objTela As Object

On Error GoTo Erro_OpcoesTela_Click

    Set objTela = Me
    
    'Se não é para atualizar a tela => sai da função
    If iAtualizaTela = 0 Then Exit Sub
    
    'Trata o evento click da combo opções
    lErro = CF("OpcoesTela_Click", objTela)
    If lErro <> SUCESSO Then gError 102933
    
    'Se Frame selecionado foi o de seleção e é para atualizar o grid
    If TabStrip1.SelectedItem.Index = TAB_RELACIONAMENTOS Then
    
        'Carrega o tab de relacionamentos
        lErro = Carrega_Tab_Relacionamentos()
        If lErro <> SUCESSO Then gError 102970
        
        iAtualizaGrid = 0
    
    End If
    
    iAlterado = 0
    iAtualizaTela = 1
    
    Exit Sub

Erro_OpcoesTela_Click:

    Select Case gErr

        Case 102933, 102970
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166628)

    End Select

End Sub

Private Sub OpcaoPadrao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoAtualizar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoAtualizar_Click

    'Se Frame selecionado foi o de seleção
    If TabStrip1.SelectedItem.Index = TAB_RELACIONAMENTOS Then
    
        'Força a atualização do Grid
        iAtualizaGrid = 1
        
        'Carrega o tab de relacionamentos
        lErro = Carrega_Tab_Relacionamentos()
        If lErro <> SUCESSO Then gError 102971
        
        iAtualizaGrid = 0
    
    End If

Exit Sub
    
Erro_BotaoAtualizar_Click:

    Select Case gErr
    
        Case 102971
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166629)
            
    End Select
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 102945
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr

        Case 102945
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166630)

    End Select

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTela As Object

On Error GoTo Erro_BotaoExcluir_Click

    Set objTela = Me
    
    lErro = CF("OpcoesTelas_Exclui", objTela)
    If lErro <> SUCESSO Then gError 102946
    
    Call Limpa_RelacionamentoClientesCons

    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 102946
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166631)

    End Select

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 102832

    'Limpa a Tela
    Call Limpa_RelacionamentoClientesCons
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 102832

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166632)

    End Select

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameTab(TabStrip1.SelectedItem.Index).Visible = True
        FrameTab(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index

        'Se Frame selecionado foi o de seleção
        If TabStrip1.SelectedItem.Index = TAB_RELACIONAMENTOS Then
        
            'Carrega o tab de relacionamentos
            lErro = Carrega_Tab_Relacionamentos()
            If lErro <> SUCESSO Then gError 102882
            
            iAtualizaGrid = 0
        
        End If
        
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 102882
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166633)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoDe_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection

    giEventoCodigo = EVENTO_CODIGODE
    
    objRelacionamentoCli.lCodigo = StrParaDbl(Codigo.Text)
    
    colSelecao.Add giFilialEmpresa
    
    Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, objEventoCodigo)
    
End Sub

Private Sub LabelCodigoAte_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection

    giEventoCodigo = EVENTO_CODIGOATE
    
    objRelacionamentoCli.lCodigo = StrParaDbl(Codigo.Text)
    
    colSelecao.Add giFilialEmpresa
    
    Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, objEventoCodigo)
    
End Sub

Private Sub UpDownPeriodoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(PeriodoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 102806

    Exit Sub

Erro_UpDownPeriodoDe:

    Select Case gErr

        Case 102806

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166634)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownPeriodoDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(PeriodoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 102794

    Exit Sub

Erro_UpDownPeriodoDe_DownClick:

    Select Case gErr

        Case 102794

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166635)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(PeriodoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 102795

    Exit Sub

Erro_UpDownPeriodoAte_UpClick:

    Select Case gErr

        Case 102795

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166636)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownPeriodoAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(PeriodoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 102807

    Exit Sub

Erro_UpDownPeriodoAte_DownClick:

    Select Case gErr

        Case 102807

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166637)

    End Select

    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelCliente_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(Cliente.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(Cliente.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = Cliente.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, "", sOrdenacao)

    Exit Sub
    
Erro_LabelCliente_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166638)
    
    End Select
    
End Sub

Private Sub AtendenteDe_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub AtendenteAte_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub TipoTodos_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
    Tipo.Text = ""
End Sub
Private Sub TipoApenas_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
    'Tipo.SetFocus
End Sub
Private Sub Tipo_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
    TipoApenas.Value = True
End Sub
Private Sub Origem_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub StatusPendente_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub StatusEncerrado_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub StatusTodos_Click()
    iAtualizaGrid = 1
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub OpcoesTela_Change()
    iAtualizaTela = 1
    iAtualizaGrid = 1
End Sub

Private Sub CodigoDe_Change()
    iAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1
End Sub
Private Sub CodigoAte_Change()
    iAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1
End Sub
Private Sub PeriodoDe_Change()
    iAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1
End Sub
Private Sub PeriodoAte_Change()
    iAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1
End Sub
Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1

    Call Cliente_Preenche

End Sub
Private Sub FilialCliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1
End Sub
Private Sub AtendenteDe_Change()
    iAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1
End Sub
Private Sub AtendenteAte_Change()
    iAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1
End Sub
Private Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
    iAtualizaGrid = 1
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO ***
Private Sub OpcoesTela_Validate(Cancel As Boolean)
    'Se a opção não foi selecionada na combo => chama a função OpcoesTela_Click
    If OpcoesTela.ListIndex = -1 Then Call OpcoesTela_Click
End Sub

Public Sub CodigoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_CodigoDe_Validate

    'Se o código não foi preenchido => sai da função
    If Len(Trim(CodigoDe.Text)) = 0 Then Exit Sub
    
    'Guarda no obj o código do relacionamento e da filialempresa
    objRelacionamentoClientes.lCodigo = StrParaLong(CodigoDe.Text)
    objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa
    
    'Lê o relacionamento no bd
    lErro = CF("RelacionamentoClientes_Le", objRelacionamentoClientes)
    If lErro <> SUCESSO And lErro <> 102508 Then gError 102807
    
    'Se não encontrou o relacionamento => erro
    If lErro = 102508 Then gError 102808
    
    'Se o código até estiver preenchido e o código de for maior que o código até => erro
    If StrParaLong(CodigoAte.Text) > 0 And StrParaLong(CodigoDe.Text) > StrParaLong(CodigoAte.Text) Then gError 102809
    
    Exit Sub
    
Erro_CodigoDe_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 102807
        
        Case 102808
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
        
        Case 102809
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166639)
    
    End Select
    
End Sub

Public Sub CodigoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_CodigoAte_Validate

    'Se o código não foi preenchido => sai da função
    If Len(Trim(CodigoAte.Text)) = 0 Then Exit Sub
    
    'Guarda no obj o código e a filial do relacionamento
    objRelacionamentoClientes.lCodigo = StrParaLong(CodigoAte.Text)
    objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa
    
    'Lê o relacionamento no bd
    lErro = CF("RelacionamentoClientes_Le", objRelacionamentoClientes)
    If lErro <> SUCESSO And lErro <> 102508 Then gError 102810
    
    'Se não encontrou o relacionamento => erro
    If lErro = 102508 Then gError 102811
    
    'Se o código até estiver preenchido e o código de for maior que o código até => erro
    If Len(Trim(CodigoDe.Text)) = 0 And StrParaLong(CodigoDe.Text) > StrParaLong(CodigoAte.Text) Then gError 102812
    
    Exit Sub
    
Erro_CodigoAte_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 102810
        
        Case 102811
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
        
        Case 102812
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166640)
    
    End Select
    
End Sub

Public Sub PeriodoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lCliente As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_PeriodoDe_Validate

    'Se a data não foi preenchida => sai da função
    If Len(Trim(PeriodoDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(PeriodoDe.Text)
    If lErro <> SUCESSO Then gError 102796
    
    'Se a data PeriodoAte está preenchida
    If StrParaDate(PeriodoAte.Text) <> DATA_NULA Then
    
        'Se a data PeriodoDe for maior que a data PeriodoAte => erro
        If StrParaDate(PeriodoDe.Text) > StrParaDate(PeriodoAte.Text) Then gError 102797
    
    End If

    Exit Sub
    
Erro_PeriodoDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 102796
        
        Case 102797
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166641)
        
    End Select

End Sub

Public Sub PeriodoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lCliente As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_PeriodoAte_Validate

    'Se a data não foi preenchida => sai da função
    If Len(Trim(PeriodoAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(PeriodoAte.Text)
    If lErro <> SUCESSO Then gError 102798
    
    'Se a data PeriodoAte está preenchida
    If StrParaDate(PeriodoDe.Text) <> DATA_NULA Then
    
        'Se a data PeriodoDe for maior que a data PeriodoAte => erro
        If StrParaDate(PeriodoDe.Text) > StrParaDate(PeriodoAte.Text) Then gError 102799
    
    End If

    Exit Sub
    
Erro_PeriodoAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 102798
        
        Case 102799
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MAIOR_DATAATE", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166642)
        
    End Select

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cliente_Validate

    'Faz a validação do cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102800
    
    Exit Sub
    
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 102800
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166643)

    End Select

End Sub

Public Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FilialCliente_Validate

    'Faz a validação da filial do cliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 102805
    
    Exit Sub
    
Erro_FilialCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 102805
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166644)

    End Select

End Sub

Public Sub AtendenteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteDe_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteDe)
    If lErro <> SUCESSO Then gError 102831
    
    'Se o atendente até foi preenchido e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 102828
    
    Exit Sub

Erro_AtendenteDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102831
        
        Case 102828
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166645)

    End Select

End Sub

Public Sub AtendenteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteAte_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteAte)
    If lErro <> SUCESSO Then gError 102820
    
    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 102829
    
    Exit Sub

Erro_AtendenteAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102820
        
        Case 102829
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166646)

    End Select

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo, "AVISO_CRIAR_TIPORELACIONAMENTOCLIENTES")
    If lErro <> SUCESSO Then gError 102815
    
    'Se a opção "Apenas" está marcada e o tipo não foi preenchido => erro
    If TipoApenas.Value = True And Len(Trim(Tipo.Text)) = 0 Then TipoTodos.Value = True
    
    Exit Sub

Erro_Tipo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102815
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166647)

    End Select

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** TRATAMENTOS DO GRID - INÍCIO ***
'ATENÇÃO: os tratamentos do grid estão incompletos, pois o grid não é editável
Public Sub GridRelacionamentos_RowColChange()
    Call Grid_RowColChange(objGridRelacionamentos)
    
    If Not (gcolRelacionamentos Is Nothing) Then
    
        'Atualiza o campo assunto
        If GridRelacionamentos.Row > 0 Then Assunto.Caption = gcolRelacionamentos(GridRelacionamentos.Row).sAssunto1 & gcolRelacionamentos(GridRelacionamentos.Row).sAssunto2
    
    End If
End Sub

Private Sub GridRelacionamentos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
        
        'Seta objTela como a Tela de Baixas a Receber
        Set PopUpMenuRelacClientes.objTela = Me
        
        'Chama o Menu PopUp
        PopupMenu PopUpMenuRelacClientes.mnuRelacionamentosClientes, vbPopupMenuRightButton
        
        'Limpa o objTela
        Set PopUpMenuRelacClientes.objTela = Nothing
        
    End If

End Sub
'*** TRATAMENTOS DO GRID - FIM ***
'*** TRATAMENTO DOS CONTROLES DA TELA - FIM ****

'*** TRATAMENTO DO EVENTO KEYDOWN  - INÍCIO ***
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodigoDe Then
            Call LabelCodigoDe_Click
        ElseIf Me.ActiveControl Is CodigoAte Then
            Call LabelCodigoAte_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
    
    ElseIf KeyCode = vbKeyF5 Then
        Call mnuRelacClientes_NovoOrcamento_Click
        
    ElseIf KeyCode = vbKeyF6 Then
        Call mnuRelacClientes_NovoPedido_Click
        
    ElseIf KeyCode = vbKeyF7 Then
        Call mnuRelacClientes_NovoRelacionamento_Click
    
    ElseIf KeyCode = vbKeyF8 Then
        Call mnuRelacClientes_Consultas_Click
        
    ElseIf KeyCode = vbKeyF9 Then
        Call mnuRelacClientes_EditarRelacionamento_Click
        
    End If

End Sub
'*** TRATAMENTO DO EVENTO KEYDOWN  - FIM ***

'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***
Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objRelacionamentoCli As New ClassRelacClientes
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objRelacionamentoCli = obj1
    
    'Se o evento foi disparado pelo LabelCodigoDe
    If giEventoCodigo = EVENTO_CODIGODE Then
        
        'Exibe o CodigoDe selecionado
        CodigoDe.PromptInclude = False
        CodigoDe.Text = objRelacionamentoCli.lCodigo
        CodigoDe.PromptInclude = True
        
        'Valida o código selecionado
        Call CodigoDe_Validate(bSGECancelDummy)
    
    'Senão, se foi disparado pelo LabelCodigoAte
    ElseIf giEventoCodigo = EVENTO_CODIGOATE Then
    
        'Exibe o CodigoAte selecionado
        CodigoAte.PromptInclude = False
        CodigoAte.Text = objRelacionamentoCli.lCodigo
        CodigoAte.PromptInclude = True
    
        'Valida o código selecionado
        Call CodigoAte_Validate(bSGECancelDummy)
    
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166648)
    
    End Select

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102813

    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr
    
        Case 102813
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166649)
    
    End Select

End Sub
'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***

'*** FUNÇÕES DE APOIO À TELA - INÍCIO ***
Private Function Inicializa_GridRelacionamentos(objGridInt As AdmGrid) As Long
'Inicializa o grid de relacionamentos

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridRelacionamentos

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Contato")
    objGridInt.colColuna.Add ("Telefone")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Origem")
    objGridInt.colColuna.Add ("Atendente")
    objGridInt.colColuna.Add ("Status")

    'campos de edição do grid
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (ClienteGrid.Name)
    objGridInt.colCampo.Add (FilialClienteGrid.Name)
    objGridInt.colCampo.Add (Contato.Name)
    objGridInt.colCampo.Add (Fone.Name)
    objGridInt.colCampo.Add (TipoRelacionamento.Name)
    objGridInt.colCampo.Add (OrigemGrid.Name)
    objGridInt.colCampo.Add (Atendente.Name)
    objGridInt.colCampo.Add (Status.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Data_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_FilialCliente_Col = 3
    iGrid_Contato_Col = 4
    iGrid_Telefone_Col = 5
    iGrid_TipoRelacionamento_Col = 6
    iGrid_Origem_Col = 7
    iGrid_Atendente_Col = 8
    iGrid_Status_Col = 9

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRelacionamentos

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_RELACIONAMENTOS

    'Largura da primeira coluna
    GridRelacionamentos.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'habilita a rotina grid enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridRelacionamentos = SUCESSO

    Exit Function

Erro_Inicializa_GridRelacionamentos:

    Inicializa_GridRelacionamentos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166650)

    End Select

    Exit Function

End Function

Private Sub Limpa_RelacionamentoClientesCons()

Dim iIndice As Integer

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    'Limpa o grid
    Call Grid_Limpa(objGridRelacionamentos)
    
    'Limpa a combo opções.
    OpcoesTela.Text = ""
    
    'Desmarca o campo padrão
    OpcaoPadrao.Value = vbUnchecked
    
    'Limpa a combo tipo
    Tipo.ListIndex = -1
    
    'Limpa a combo filial
    FilialCliente.Clear
    
    'Preenche valores default da tela
    'Data Inicial
    PeriodoDe.PromptInclude = False
    PeriodoDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    PeriodoDe.PromptInclude = True
    
    'Data Final
    PeriodoAte.PromptInclude = False
    PeriodoAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    PeriodoAte.PromptInclude = True
    
    'Limpa as combos de atendentes, antes de selecionar o atendente padrão
    AtendenteDe.Text = ""
    AtendenteAte.Text = ""
    
    'Atendentes
    'Para cada atendente da combo AtendenteDe
    For iIndice = 0 To AtendenteDe.ListCount - 1
    
        'Se o conteúdo do atendente for igual ao seu código + "-" + nome reduzido do usuário ativo
        If AtendenteDe.List(iIndice) = AtendenteDe.ItemData(iIndice) & SEPARADOR & gsUsuario Then
        
            'Significa que achou o atendente "default"
            'Seleciona o atendente na combo AtendenteDe
            AtendenteDe.ListIndex = iIndice
            
            'Seleciona o atendente na combo AtendenteAt
            AtendenteAte.ListIndex = iIndice
            
            'Sai do For
            Exit For
        End If
    Next
    
    'TipoTodos
    TipoTodos.Value = True
    
    'Origem
    Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
    
    'Status
    StatusPendente.Value = True
    
    'Limpa o campo assunto
    Assunto.Caption = ""
    
    iAlterado = 0
    iClienteAlterado = 0
    
End Sub

Private Function Valida_Cliente() As Long
'Faz a validação do cliente

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Valida_Cliente

    'Se o campo cliente não foi alterado => sai da função
    If iClienteAlterado = 0 Then Exit Function

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 102801

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 102802

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialCliente, iCodFilial)
        
    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialCliente.Clear
        
    End If
    
    iClienteAlterado = 0
    
    Valida_Cliente = SUCESSO

    Exit Function

Erro_Valida_Cliente:

    Valida_Cliente = gErr
    
    Select Case gErr

        Case 102801, 102802
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166651)

    End Select

    Exit Function

End Function

Private Function Valida_FilialCliente() As Long
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer

On Error GoTo Erro_Valida_FilialCliente

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialCliente.Text)) = 0 Then Exit Function

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(FilialCliente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102803

    'Se foi digitado o nome da filial
    'e esse nome não foi encontrado na combo => erro
    If lErro = 6731 Then gError 102804
    
    Valida_FilialCliente = SUCESSO
    
    Exit Function

Erro_Valida_FilialCliente:

    Valida_FilialCliente = gErr

    Select Case gErr

        Case 102803
        
        Case 102804

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166652)

    End Select

    Exit Function

End Function

Private Function Carrega_Tab_Relacionamentos() As Long
'Dispara as funções que fazem a carga do tab relacionamentos

Dim lErro As Long
Dim objRelacionamentoClientesCons As New ClassRelacClientesCons

On Error GoTo Erro_Carrega_Tab_Relacionamentos

    'Se não houve alteração na tela => sai da função
    If iAtualizaGrid = 0 Then Exit Function
    
    'Limpa o Grid
    Call Grid_Limpa(objGridRelacionamentos)
    
    'Limpa o assunto
    Assunto.Caption = ""
    
    'Move os dados do tab seleção para memória
    lErro = Move_TabSelecao_Memoria(objRelacionamentoClientesCons)
    If lErro <> SUCESSO Then gError 102872
    
    'Lê os relacionamentos com os filtros passados
    lErro = CF("RelacionamentoClientesCons_Le", objRelacionamentoClientesCons)
    If lErro <> SUCESSO And lErro <> 102871 Then gError 102873
    
    'Se não encontrou nenhum relacionamento
    If lErro = 102871 Then
    
        'Limpa a coleção global de relacionamentos
        Set gcolRelacionamentos = Nothing
        
        'sai por erro
        gError 102874
    End If
    
    'Carrega o grid com os relacionamentos encontrados
    lErro = Carrega_GridRelacionamentos(objRelacionamentoClientesCons.colRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102875

    Carrega_Tab_Relacionamentos = SUCESSO

    Exit Function

Erro_Carrega_Tab_Relacionamentos:

    Carrega_Tab_Relacionamentos = gErr

    Select Case gErr

        Case 102872, 102873, 102875
        
        Case 102874
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO1", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166653)

    End Select

End Function

Private Function Move_TabSelecao_Memoria(ByVal objRelacionamentoClientesCons As ClassRelacClientesCons) As Long
'Move para a memória os dados do tab seleção

Dim lErro As Long
Dim lCliente As Long
Dim iFilialCliente As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    'FilialEmpresa
    objRelacionamentoClientesCons.iFilialEmpresa = giFilialEmpresa
    
    'Código De/Até
    objRelacionamentoClientesCons.lCodigoDe = StrParaLong(CodigoDe.Text)
    objRelacionamentoClientesCons.lCodigoAte = StrParaLong(CodigoAte.Text)
    
    'Data De/Até
    objRelacionamentoClientesCons.dtDataDe = StrParaDate(PeriodoDe.Text)
    objRelacionamentoClientesCons.dtDataAte = StrParaDate(PeriodoAte.Text)
    
    'Cliente / FilialCliente
    lErro = Obtem_CodCliente(lCliente, iFilialCliente)
    If lErro <> SUCESSO Then gError 102835
    
    objRelacionamentoClientesCons.lCliente = lCliente
    objRelacionamentoClientesCons.iFilialCliente = iFilialCliente
    
    'Atendente De/Até
    objRelacionamentoClientesCons.iAtendenteDe = Codigo_Extrai(AtendenteDe.Text)
    objRelacionamentoClientesCons.iAtendenteAte = Codigo_Extrai(AtendenteAte.Text)
    
    'Tipo
    objRelacionamentoClientesCons.lTipo = Codigo_Extrai(Tipo.Text)
    
    'Origem
    If Origem.Text = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO Then
        objRelacionamentoClientesCons.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE
    ElseIf Origem.Text = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO Then
        objRelacionamentoClientesCons.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA
    End If
    
    'Status
    If StatusEncerrado.Value = True Then
        objRelacionamentoClientesCons.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO
    ElseIf StatusTodos.Value = True Then
        objRelacionamentoClientesCons.iStatus = -1
    End If
    
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case 102835
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166654)

    End Select

End Function

Private Function Obtem_CodCliente(lCliente As Long, iFilialCliente As Integer) As Long
'Obtém o código do cliente e da filial que estão na tela e guarda-os no objClienteContatos

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_Obtem_CodCliente

    'Se o cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
    
        '*** Leitura do cliente a partir do nome reduzido para obter o seu código ***
        
        'Guarda o nome reduzido do cliente
        objcliente.sNomeReduzido = Trim(Cliente.Text)
        
        'Faz a leitura do cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 102833
        
        'Se não encontrou o cliente => erro
        If lErro = 12348 Then gError 102834
        
        'Devolve o código do cliente
        lCliente = objcliente.lCodigo
        '*** Fim da leitura de cliente ***
        
        'Devolve o código da filial do cliente
        iFilialCliente = Codigo_Extrai(FilialCliente.Text)
        
    End If

    Obtem_CodCliente = SUCESSO

    Exit Function

Erro_Obtem_CodCliente:

    Obtem_CodCliente = gErr

    Select Case gErr

        Case 102833

        Case 102834
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166655)

    End Select

End Function

Private Function Carrega_GridRelacionamentos(ByVal colRelacionamentos As Collection) As Long
'Carrega o grid com os relacionamento da coleção

Dim lErro As Long
Dim objRelacionamentoClientes As ClassRelacClientes
Dim iLinha As Integer
Dim objAtendente As New ClassAtendentes
Dim objCamposGenericosValores As New ClassCamposGenericosValores
Dim objcliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Carrega_GridRelacionamentos

    'Agiliza a exibição da tela
    DoEvents
    
    'Para cada relacionamento na coleção
    For Each objRelacionamentoClientes In colRelacionamentos
    
        iLinha = iLinha + 1
        
        'data
        GridRelacionamentos.TextMatrix(iLinha, iGrid_Data_Col) = objRelacionamentoClientes.dtData
        
        'Origem
        If objRelacionamentoClientes.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE Then
            GridRelacionamentos.TextMatrix(iLinha, iGrid_Origem_Col) = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO
        Else
            GridRelacionamentos.TextMatrix(iLinha, iGrid_Origem_Col) = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO
        End If
        
        '*** ATENDENTE - INÍCIO ***
        'Guarda no obj o código e a filial do atendente
        objAtendente.iCodigo = objRelacionamentoClientes.iAtendente
        objAtendente.iFilialEmpresa = objRelacionamentoClientes.iFilialEmpresa
        
        'Lê os dados do atendente para obter o nome reduzido
        lErro = CF("Atendentes_Le", objAtendente)
        If lErro <> SUCESSO And lErro <> 102752 Then gError 102883
    
        'Se não encontrou o atendente => erro
        If lErro = 102752 Then gError 102884
        
        'Exibe o atendente no grid
        GridRelacionamentos.TextMatrix(iLinha, iGrid_Atendente_Col) = objAtendente.sNomeReduzido
        '*** ATENDENTE - FIM ***
        
        '*** TIPO DE RELACIONAMENTO - INÍCIO ***
        'Guarda no obj o código do campo e do tipo a ser lido
        objCamposGenericosValores.lCodCampo = CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES
        objCamposGenericosValores.lCodValor = objRelacionamentoClientes.lTipo
        
        'Lê os dados do tipo para obter a descrição
        lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCamposGenericosValores)
        If lErro <> SUCESSO And lErro <> 102399 Then gError 102885
        
        'Se não encontrou => erro
        If lErro = 102399 Then gError 102886
        
        'Exibe o tipo de relacionamento no grid
        GridRelacionamentos.TextMatrix(iLinha, iGrid_TipoRelacionamento_Col) = objCamposGenericosValores.sValor
        '*** TIPO DE RELACIONAMENTO - FIM ***
        
        '*** CLIENTE - INÍCIO ***
        'Guarda o código do cliente no obj
        objcliente.lCodigo = objRelacionamentoClientes.lCliente
        
        'Lê o cliente para obter o nome reduzido
        lErro = CF("Cliente_Le", objcliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 102876
        
        'Se não encontrou o cliente => erro
        If lErro = 12293 Then gError 102879
        
        'Exibe o nome do cliente no grid
        GridRelacionamentos.TextMatrix(iLinha, iGrid_Cliente_Col) = objcliente.sNomeReduzido
        '*** CLIENTE - FIM ***
        
        '*** FILIALCLIENTE - INÍCIO ***
        'Guarda o código do cliente e da filial no obj
        objFilialCliente.lCodCliente = objRelacionamentoClientes.lCliente
        objFilialCliente.iCodFilial = objRelacionamentoClientes.iFilialCliente
        
        'Lê a filial do cliente para obter o nome reduzido
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 102877
    
        'Se não existe a filial na tabela FiliaisClientes
        If lErro = 12567 Then gError 102880
        
        'Exibe o nome da filial no grid
        GridRelacionamentos.TextMatrix(iLinha, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
        '*** FILIALCLIENTE - FIM ***
        
        '*** CONTATO - INÍCIO ***
        'Se o relacionamento possui um contato
        If objRelacionamentoClientes.iContato > 0 Then
        
            'Guarda no obj o código do cliente, da filial e do contato
            objClienteContatos.lCliente = objRelacionamentoClientes.lCliente
            objClienteContatos.iFilialCliente = objRelacionamentoClientes.iFilialCliente
            objClienteContatos.iCodigo = objRelacionamentoClientes.iContato
            
            'Lê o contato para obter o nome e o telefone
            lErro = CF("ClienteContatos_Le", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 102653 Then gError 102878
            
            'Se não encontrou o contato => erro
            If lErro = 102653 Then gError 102881
            
            'Exibe o nome do contato no grid
            GridRelacionamentos.TextMatrix(iLinha, iGrid_Contato_Col) = objClienteContatos.sContato
           
            'Exibe o telefone no grid
            GridRelacionamentos.TextMatrix(iLinha, iGrid_Telefone_Col) = objClienteContatos.sTelefone
        
        End If
        '*** CONTATO - FIM ***
        
        'Status
        If objRelacionamentoClientes.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO Then
            GridRelacionamentos.TextMatrix(iLinha, iGrid_Status_Col) = "Encerrado"
        Else
            GridRelacionamentos.TextMatrix(iLinha, iGrid_Status_Col) = "Pendente"
        End If
        
    Next

    'Atualiza o número de linhas no grid
    objGridRelacionamentos.iLinhasExistentes = iLinha - 1
    
    'Faz a coleção global de relacionamento apontar para a coleção que foi
    'utilizada para atualizar o grid
    Set gcolRelacionamentos = colRelacionamentos
    
    Carrega_GridRelacionamentos = SUCESSO

    Exit Function

Erro_Carrega_GridRelacionamentos:

    Call Grid_Limpa(objGridRelacionamentos)
    
    Carrega_GridRelacionamentos = gErr

    Select Case gErr

        Case 102876 To 102878, 102883, 102885
        
        Case 102879
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objRelacionamentoClientes.lCliente)
            
        Case 102880
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objRelacionamentoClientes.iFilialCliente, objRelacionamentoClientes.lCliente)
        
        Case 102881
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.iContato, objcliente.sNomeReduzido, objFilialCliente.sNome)
        
        Case 102884
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.iAtendente, objRelacionamentoClientes.iFilialEmpresa)
        
        Case 102886
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPORELACIONAMENTOCLI_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lTipo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166656)

    End Select

End Function

'*** GRAVAÇÃO - INÍCIO ***
Public Function Gravar_Registro() As Long

Dim objTela As Object
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    Set objTela = Me
    
    lErro = CF("OpcoesTelas_Grava", objTela)
    If lErro <> SUCESSO Then gError 102930
    
    Call Limpa_RelacionamentoClientesCons
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 102930
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166657)
        
    End Select
    
End Function
'*** GRAVAÇÃO - FIM ***

'*** FUNCIONAMENTO DO MENU POPUP DO GRID - INÍCIO ***
Public Sub mnuRelacClientes_NovoOrcamento_Click()

Dim objOrcamentoVenda As New ClassOrcamentoVenda

    'Se não foi selecionada uma linha do grid => sai da função
    If GridRelacionamentos.Row = 0 Then Exit Sub
    
    'Se não há relacionamentos na coleção global => sai da função
    If gcolRelacionamentos Is Nothing Then Exit Sub
    
    objOrcamentoVenda.lCliente = gcolRelacionamentos(GridRelacionamentos.Row).lCliente
    objOrcamentoVenda.iFilial = gcolRelacionamentos(GridRelacionamentos.Row).iFilialCliente
    
    Call Chama_Tela("OrcamentoVenda", objOrcamentoVenda)
End Sub

Public Sub mnuRelacClientes_NovoPedido_Click()
    
Dim objPedidoVenda As New ClassPedidoDeVenda

    'Se não foi selecionada uma linha do grid => sai da função
    If GridRelacionamentos.Row = 0 Then Exit Sub
    
    'Se não há relacionamentos na coleção global => sai da função
    If gcolRelacionamentos Is Nothing Then Exit Sub
    
    'Guarda no obj o código do cliente e da filial do cliente
    objPedidoVenda.lCliente = gcolRelacionamentos(GridRelacionamentos.Row).lCliente
    objPedidoVenda.iFilial = gcolRelacionamentos(GridRelacionamentos.Row).iFilialCliente
    
    'Chama a tela de Pedido de Venda
    Call Chama_Tela("PedidoVenda", objPedidoVenda)
    
End Sub

Public Sub mnuRelacClientes_Consultas_Click()

Dim objcliente As New ClassCliente

    'Se não foi selecionada uma linha do grid => sai da função
    If GridRelacionamentos.Row = 0 Then Exit Sub
    
    'Se não há relacionamentos na coleção global => sai da função
    If gcolRelacionamentos Is Nothing Then Exit Sub
    
    'Guarda no obj o código do cliente
    objcliente.lCodigo = gcolRelacionamentos(GridRelacionamentos.Row).lCliente
    
    Call Chama_Tela("ClienteConsulta", objcliente)
    
End Sub

Public Sub mnuRelacClientes_EditarRelacionamento_Click()
    
Dim objRelacionamentoClientes As New ClassRelacClientes

    'Se não foi selecionada uma linha do grid => sai da função
    If GridRelacionamentos.Row = 0 Then Exit Sub
    
    'Se não há relacionamentos na coleção global => sai da função
    If gcolRelacionamentos Is Nothing Then Exit Sub
    
    'Guarda no obj os dados do relacionamento em questão
    objRelacionamentoClientes.lCodigo = gcolRelacionamentos(GridRelacionamentos.Row).lCodigo
    objRelacionamentoClientes.iFilialEmpresa = gcolRelacionamentos(GridRelacionamentos.Row).iFilialEmpresa
    
    'Chama a tela de relacionamentos
    Call Chama_Tela("RelacionamentoClientes", objRelacionamentoClientes)

End Sub

Public Sub mnuRelacClientes_NovoRelacionamento_Click()

Dim objRelacionamentoClientes As New ClassRelacClientes

    'Se não foi selecionada uma linha do grid => sai da função
    If GridRelacionamentos.Row = 0 Then Exit Sub
    
    'Se não há relacionamentos na coleção global => sai da função
    If gcolRelacionamentos Is Nothing Then Exit Sub
    
    'Guarda no obj os dados que serão indicados como default para a criação de um novo relacionamento
    objRelacionamentoClientes.lCliente = gcolRelacionamentos(GridRelacionamentos.Row).lCliente
    objRelacionamentoClientes.iFilialCliente = gcolRelacionamentos(GridRelacionamentos.Row).iFilialCliente
    objRelacionamentoClientes.iContato = gcolRelacionamentos(GridRelacionamentos.Row).iContato
    objRelacionamentoClientes.iAtendente = gcolRelacionamentos(GridRelacionamentos.Row).iAtendente
    objRelacionamentoClientes.dtData = gdtDataAtual
    objRelacionamentoClientes.dtDataFim = DATA_NULA
    objRelacionamentoClientes.dtDataProxCobr = DATA_NULA
    objRelacionamentoClientes.dtDataPrevReceb = DATA_NULA
    
    'Chama a tela de relacionamentos
    Call Chama_Tela("RelacionamentoClientes", objRelacionamentoClientes)
End Sub
'*** FUNCIONAMENTO DO MENU POPUP DO GRID - INÍCIO ***

'*** FUNÇÕES DE APOIO À TELA - FIM ***

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - INÍCIO ***
Private Sub LabelDataDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataDe, Source, X, Y)
End Sub

Private Sub LabelDataDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataAte, Source, X, Y)
End Sub

Private Sub LabelDataAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialCliente, Source, X, Y)
End Sub

Private Sub LabelFilialCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelAtendenteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAtendenteDe, Source, X, Y)
End Sub

Private Sub LabelAtendenteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAtendenteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelAtendenteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAtendenteAte, Source, X, Y)
End Sub

Private Sub LabelAtendenteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAtendenteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrigem, Source, X, Y)
End Sub

Private Sub LabelOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrigem, Button, Shift, X, Y)
End Sub

Private Sub LabelAssunto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAssunto, Source, X, Y)
End Sub

Private Sub LabelAssunto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAssunto, Button, Shift, X, Y)
End Sub
'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - FIM ***

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134031

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134031

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166658)

    End Select
    
    Exit Sub

End Sub

