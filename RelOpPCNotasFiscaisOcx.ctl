VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPCNotasFiscaisOcx 
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   ScaleHeight     =   4305
   ScaleWidth      =   9105
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   2
      Left            =   630
      TabIndex        =   40
      Top             =   1260
      Visible         =   0   'False
      Width           =   7890
      Begin VB.Frame Frame6 
         Caption         =   "Destinatários"
         Height          =   1050
         Left            =   90
         TabIndex        =   46
         Top             =   1395
         Width           =   7665
         Begin VB.Frame Frame7 
            Caption         =   "Tipo"
            Height          =   555
            Left            =   135
            TabIndex        =   52
            Top             =   270
            Width           =   3930
            Begin VB.OptionButton TipoDestino 
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
               Index           =   0
               Left            =   105
               TabIndex        =   13
               Top             =   225
               Width           =   870
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Filial Empresa"
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
               Left            =   990
               TabIndex        =   14
               Top             =   225
               Value           =   -1  'True
               Width           =   1515
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Fornecedor"
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
               Left            =   2580
               TabIndex        =   15
               Top             =   225
               Width           =   1305
            End
         End
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   870
            Index           =   1
            Left            =   4260
            TabIndex        =   47
            Top             =   135
            Width           =   3240
            Begin VB.ComboBox FilialEmpresa 
               Height          =   315
               ItemData        =   "RelOpPCNotasFiscaisOcx.ctx":0000
               Left            =   720
               List            =   "RelOpPCNotasFiscaisOcx.ctx":0002
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   315
               Width           =   2160
            End
            Begin VB.Label LabelFilialEmpDestino 
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
               Height          =   195
               Left            =   135
               TabIndex        =   48
               Top             =   375
               Width           =   465
            End
         End
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Height          =   795
            Index           =   2
            Left            =   4185
            TabIndex        =   49
            Top             =   180
            Visible         =   0   'False
            Width           =   3285
            Begin VB.ComboBox FilialFornecedor 
               Height          =   315
               Left            =   1125
               TabIndex        =   18
               Top             =   435
               Width           =   2160
            End
            Begin MSMask.MaskEdBox Fornecedor 
               Height          =   300
               Left            =   1125
               TabIndex        =   17
               Top             =   75
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label LabelFilialFornDestino 
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
               Height          =   195
               Left            =   600
               TabIndex        =   51
               Top             =   495
               Width           =   465
            End
            Begin VB.Label LabelFornDestino 
               AutoSize        =   -1  'True
               Caption         =   "Fornecedor:"
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
               Left            =   75
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   50
               Top             =   135
               Width           =   1035
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filial Empresa"
         Height          =   1155
         Left            =   90
         TabIndex        =   41
         Top             =   135
         Width           =   7665
         Begin MSMask.MaskEdBox CodFilialDe 
            Height          =   300
            Left            =   1155
            TabIndex        =   9
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodFilialAte 
            Height          =   300
            Left            =   4875
            TabIndex        =   10
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeAte 
            Height          =   300
            Left            =   4875
            TabIndex        =   12
            Top             =   705
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeDe 
            Height          =   300
            Left            =   1140
            TabIndex        =   11
            Top             =   690
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeAte 
            AutoSize        =   -1  'True
            Caption         =   "Nome Até:"
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
            Left            =   3960
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   45
            Top             =   765
            Width           =   900
         End
         Begin VB.Label LabelCodFilialAte 
            AutoSize        =   -1  'True
            Caption         =   "Código Até:"
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
            Left            =   3855
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   44
            Top             =   345
            Width           =   1005
         End
         Begin VB.Label LabelNomeDe 
            AutoSize        =   -1  'True
            Caption         =   "Nome De:"
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
            Left            =   270
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   43
            Top             =   765
            Width           =   855
         End
         Begin VB.Label LabelCodFilialDe 
            AutoSize        =   -1  'True
            Caption         =   "Código De:"
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   42
            Top             =   323
            Width           =   960
         End
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpPCNotasFiscaisOcx.ctx":0004
      Left            =   1575
      List            =   "RelOpPCNotasFiscaisOcx.ctx":0014
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPCNotasFiscaisOcx.ctx":0050
      Left            =   1575
      List            =   "RelOpPCNotasFiscaisOcx.ctx":0052
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   3090
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4950
      Picture         =   "RelOpPCNotasFiscaisOcx.ctx":0054
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   135
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6825
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPCNotasFiscaisOcx.ctx":0156
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPCNotasFiscaisOcx.ctx":02D4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPCNotasFiscaisOcx.ctx":0806
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPCNotasFiscaisOcx.ctx":0990
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pedidos de Compra"
      Height          =   2385
      Index           =   1
      Left            =   720
      TabIndex        =   28
      Top             =   1395
      Width           =   7725
      Begin VB.Frame FrameCodigo 
         Caption         =   "Número"
         Height          =   675
         Left            =   180
         TabIndex        =   37
         Top             =   375
         Width           =   3270
         Begin MSMask.MaskEdBox NumeroDe 
            Height          =   300
            Left            =   525
            TabIndex        =   2
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroAte 
            Height          =   300
            Left            =   2145
            TabIndex        =   3
            Top             =   240
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelNumeroDe 
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
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   315
            Width           =   315
         End
         Begin VB.Label LabelNumeroAte 
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
            Left            =   1695
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   300
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data de Envio"
         Height          =   690
         Left            =   165
         TabIndex        =   32
         Top             =   1320
         Width           =   3900
         Begin MSComCtl2.UpDown UpDownDataDe 
            Height          =   315
            Left            =   1665
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDe 
            Height          =   315
            Left            =   480
            TabIndex        =   6
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAte 
            Height          =   315
            Left            =   3540
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   315
            Left            =   2355
            TabIndex        =   7
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
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
            Left            =   165
            TabIndex        =   36
            Top             =   315
            Width           =   315
         End
         Begin VB.Label Label3 
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
            Left            =   1980
            TabIndex        =   35
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fornecedores"
         Height          =   705
         Left            =   3630
         TabIndex        =   29
         Top             =   360
         Width           =   3900
         Begin MSMask.MaskEdBox FornecedorDe 
            Height          =   300
            Left            =   540
            TabIndex        =   4
            Top             =   255
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorAte 
            Height          =   300
            Left            =   2340
            TabIndex        =   5
            Top             =   255
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelFornecedorDe 
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   31
            Top             =   315
            Width           =   315
         End
         Begin VB.Label LabelFornecedorAte 
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
            Left            =   1950
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.CheckBox CheckItens 
         Caption         =   "Exibe Item a Item"
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
         Left            =   4800
         TabIndex        =   8
         Top             =   1560
         Width           =   2070
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3075
      Left            =   540
      TabIndex        =   27
      Top             =   945
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   5424
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedido"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Continuação"
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Ordenados Por:"
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
      Left            =   225
      TabIndex        =   26
      Top             =   510
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   25
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPCNotasFiscaisOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpPCNotasFiscais
Const ORD_POR_CODIGO = 0
Const ORD_POR_DATA = 1
Const ORD_POR_FORNECEDOR = 2
Const ORD_POR_COMPRADOR = 3

Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNumeroDe As AdmEvento
Attribute objEventoNumeroDe.VB_VarHelpID = -1
Private WithEvents objEventoNumeroAte As AdmEvento
Attribute objEventoNumeroAte.VB_VarHelpID = -1
Private WithEvents objEventoFornDe As AdmEvento
Attribute objEventoFornDe.VB_VarHelpID = -1
Private WithEvents objEventoFornAte As AdmEvento
Attribute objEventoFornAte.VB_VarHelpID = -1
Private WithEvents objEventoCompradorDe As AdmEvento
Attribute objEventoCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoCompradorAte As AdmEvento
Attribute objEventoCompradorAte.VB_VarHelpID = -1
Private WithEvents objEventoFornDestino As AdmEvento
Attribute objEventoFornDestino.VB_VarHelpID = -1

Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giTipoDestinoAtual As Integer


Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 73591

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 73592

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 73591

        Case 73592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170720)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 73593

    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckItens.Value = vbUnchecked
    FilialEmpresa.ListIndex = 0
    
    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 73593

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170721)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoFornDestino = New AdmEvento
    Set objEventoCodPCDe = New AdmEvento
    Set objEventoCodPCAte = New AdmEvento
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
    Set objEventoFornDe = New AdmEvento
    Set objEventoFornAte = New AdmEvento
    Set objEventoNumeroDe = New AdmEvento
    Set objEventoNumeroAte = New AdmEvento
    Set objEventoCompradorDe = New AdmEvento
    Set objEventoCompradorAte = New AdmEvento
    
    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 73685

    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
        FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
    Next

    giTipoDestinoAtual = 1
    FilialEmpresa.ListIndex = 0
    
    iFrameAtual = 1
    
    ComboOrdenacao.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 73685
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170722)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoFornDestino = Nothing
    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
    Set objEventoNumeroDe = Nothing
    Set objEventoNumeroAte = Nothing
    Set objEventoFornDe = Nothing
    Set objEventoFornAte = Nothing
    Set objEventoCompradorDe = Nothing
    Set objEventoCompradorAte = Nothing
    
End Sub


Private Sub CodFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialAte, iAlterado)
    
End Sub

Private Sub CodFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialDe, iAlterado)
    
End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    
End Sub

Private Sub FilialFornecedor_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialFornecedor_Click

    If FilialFornecedor.ListIndex = -1 Then Exit Sub
    
    objFilialFornecedor.iCodFilial = FilialFornecedor.ItemData(FilialFornecedor.ListIndex)

    'Busca no BD a FilialFornecedor
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 73683

    If lErro = 18272 Then gError 73684

    Exit Sub

Erro_FilialFornecedor_Click:

    Select Case gErr

        Case 73683

        Case 73684
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, FilialFornecedor.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170723)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorAte, iAlterado)
    
End Sub

Private Sub FornecedorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorDe, iAlterado)
    
End Sub

Private Sub LabelFornDestino_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    objFornecedor.sNomeReduzido = Fornecedor.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornDestino)
    Exit Sub
    
End Sub

Private Sub NumeroAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroAte, iAlterado)
    
End Sub

Private Sub NumeroDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroDe, iAlterado)
    
End Sub

Private Sub LabelFornecedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorDe_Click

    If Len(Trim(FornecedorDe.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorDe.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornDe)
    
    Exit Sub

Erro_LabelFornecedorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170724)

    End Select

    Exit Sub
    
End Sub
Private Sub LabelFornecedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorAte_Click

    If Len(Trim(FornecedorAte.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornAte)
    
    Exit Sub

Erro_LabelFornecedorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170725)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelNumeroAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompras As New ClassPedidoCompras

On Error GoTo Erro_LabelNumeroAte_Click

    If Len(Trim(NumeroAte.Text)) > 0 Then
        'Preenche com o numero da tela
        objPedCompras.lCodigo = StrParaLong(NumeroAte.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompras, objEventoNumeroAte)

   Exit Sub

Erro_LabelNumeroAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170726)

    End Select

    Exit Sub

End Sub
Private Sub LabelNumeroDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelNumeroDe_Click

    If Len(Trim(NumeroDe.Text)) > 0 Then
        'Preenche com o numero da tela
        objPedCompra.lCodigo = StrParaLong(NumeroDe.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoNumeroDe)

   Exit Sub

Erro_LabelNumeroDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170727)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 73595

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73595
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170728)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 73596

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73596
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170729)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornDestino_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    Fornecedor.Text = objFornecedor.sNomeReduzido
    'Dispara Validate de Fornecedor
    bCancel = False
    Fornecedor_Validate (bCancel)

    Me.Show

End Sub

Private Sub TabStrip1_Click()

    'Se frame atual corresponde ao tab selecionado, sai da rotina
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True

    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False

    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index


End Sub

Private Sub TipoDestino_Click(Index As Integer)

Dim lErro As Long

On Error GoTo Erro_TipoDestino_Click


If Index = giTipoDestinoAtual Then Exit Sub

    If Index <> 0 Then
        
        FilialFornecedor.Enabled = True
        Fornecedor.Enabled = True
        LabelFornDestino.Enabled = True
        LabelFilialFornDestino.Enabled = True
        FilialEmpresa.Enabled = True
        LabelFilialEmpDestino.Enabled = True
        FrameTipo(1).Visible = False
        
        'Torna Frame correspondente a Index visivel
        FrameTipo(Index).Visible = True

        'Torna Frame atual invisivel
        If giTipoDestinoAtual <> 0 Then FrameTipo(giTipoDestinoAtual).Visible = False

        'Armazena novo valor de iFrameTipoDestinoAtual
        giTipoDestinoAtual = Index

        If Index <> 1 Then

            FrameTipo(Index - 1).Visible = False
            FrameTipo(Index).Visible = True
            
            'Verifica se o Fornecedor e sua Filial estão preenchidos
            If Len(Trim(Fornecedor.Text)) > 0 And Len(Trim(FilialFornecedor.Text)) > 0 Then
            
                FilialFornecedor_Click

            End If
        Else
            
            FrameTipo(Index + 1).Visible = False
            FrameTipo(Index).Visible = True
            Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)

        End If

    End If
    If Index = 0 Then
    
        FilialEmpresa.Enabled = False
        FilialEmpresa.ListIndex = -1
        LabelFilialEmpDestino.Enabled = False
        Fornecedor.Enabled = False
        Fornecedor.Text = ""
        FilialFornecedor.Enabled = False
        FilialFornecedor.Text = ""
        LabelFornDestino.Enabled = False
        LabelFilialFornDestino.Enabled = False
        giTipoDestinoAtual = Index
        
    End If
    
    Exit Sub
    
Erro_TipoDestino_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170730)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73597

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 73597
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170731)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73598

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 73598
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170732)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73599

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 73599
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170733)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73600

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 73600
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170734)

    End Select

    Exit Sub

End Sub


Private Sub LabelCodFilialDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodFilialDe_Click

    If Len(Trim(CodFilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialDe)

   Exit Sub

Erro_LabelCodFilialDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170735)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodFilialAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodFilialAte_Click

    If Len(Trim(CodFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialAte)

   Exit Sub

Erro_LabelCodFilialAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170736)

    End Select

    Exit Sub

End Sub




Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objFilialEmpresa.sNome = NomeDe.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170737)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.sNome = NomeAte.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170738)

    End Select

    Exit Sub

End Sub
Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor
Dim lCodigo As Long

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se Fornec esta preenchido
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        'Le Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 73674

        'Le as filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6698 Then gError 73675

        'Preenche a combo de filiais
        Call CF("Filial_Preenche", FilialFornecedor, colCodigoNome)

        'Seleciona a filial na combo de filiais
        Call CF("Filial_Seleciona", FilialFornecedor, iCodFilial)

    Else
        'Limpa a combobox
        FilialFornecedor.Clear

    End If
  
    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 73674, 73675

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170739)

    End Select

    Exit Sub

End Sub
Private Sub FilialFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult
Dim sNomeRed As String
Dim objEndereco As New ClassEndereco
Dim objPais As New ClassPais

On Error GoTo Erro_FilialFornecedor_Validate

    'Verifica se FilialFornecedor esta preenchida
    If Len(Trim(FilialFornecedor.Text)) > 0 Then

        'Verifica se FilialFornecedor esta selecionada
        If FilialFornecedor.ListIndex <> -1 Then Exit Sub

        'Seleciona combo box de FilialFornecedor
        lErro = Combo_Seleciona(FilialFornecedor, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 73676

        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then

            'Verifica de o fornecedor foi digitado
            If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 73677

            sNomeRed = Fornecedor.Text

            objFilialFornecedor.iCodFilial = iCodigo

            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 73678

            If lErro = 18272 Then gError 73679

            'Coloca na tela Codigo e Nome Reduzido de FilialFornec
            FilialFornecedor.Text = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

        End If

        'Não encontrou valor informado que era STRING
        If lErro = 6731 Then gError 73680

    End If

    Exit Sub

Erro_FilialFornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 73677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 73678

        Case 73679, 73680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, objFilialFornecedor.sNome)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170740)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeDe.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeAte.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorDe.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorAte.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub



Private Sub objEventoNumeroAte_evSelecao(obj1 As Object)

Dim objPedidoCompra As New ClassPedidoCompras

    Set objPedidoCompra = obj1

    NumeroAte.Text = CStr(objPedidoCompra.lCodigo)

    Me.Show

End Sub

Private Sub objEventoNumeroDe_evSelecao(obj1 As Object)

Dim objPedidoCompra As New ClassPedidoCompras

    Set objPedidoCompra = obj1

    NumeroDe.Text = CStr(objPedidoCompra.lCodigo)

    Me.Show

End Sub


Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 73602

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73603

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 73604

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 73605

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 73602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 73603 To 73605

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170741)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 73606

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 73607

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 73606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 73607

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170742)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73608

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFSerie", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFNumero", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
            
            Case ORD_POR_DATA

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEnvio", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFSerie", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFNumero", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
                
            Case ORD_POR_FORNECEDOR
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFSerie", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFNumero", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
                
            Case ORD_POR_COMPRADOR
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CompradorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFSerie", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NFNumero", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
                
            Case Else
                gError 74955

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 73608, 74955

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170743)

    End Select

    Exit Sub

End Sub


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodFilial_I As String
Dim sCodFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sComp_I As String
Dim sComp_F As String
Dim sNumero_I As String
Dim sNumero_F As String
Dim sForn_I As String
Dim sForn_F As String
Dim sCheck As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sNumero_I, sNumero_F, sForn_I, sForn_F, sComp_I, sComp_F)
    If lErro <> SUCESSO Then gError 73609

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 73610

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 73611

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73612

    lErro = objRelOpcoes.IncluirParametro("NCODPCINIC", sNumero_I)
    If lErro <> AD_BOOL_TRUE Then gError 73613

    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sForn_I)
    If lErro <> AD_BOOL_TRUE Then gError 73614
    
    lErro = objRelOpcoes.IncluirParametro("NCODCOMPINIC", sComp_I)
    If lErro <> AD_BOOL_TRUE Then gError 73660
    
    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAENVINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAENVINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73616

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 73617

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73618

    lErro = objRelOpcoes.IncluirParametro("NCODPCFIM", sNumero_F)
    If lErro <> AD_BOOL_TRUE Then gError 73619

    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sForn_F)
    If lErro <> AD_BOOL_TRUE Then gError 73620

    lErro = objRelOpcoes.IncluirParametro("NCODCOMPFIM", sComp_F)
    If lErro <> AD_BOOL_TRUE Then gError 73659

    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAENVFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAENVFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73622

    'Exibe Itens
    If CheckItens.Value = 0 Then
        sCheck = 0
        gobjRelatorio.sNomeTsk = "pcxnf"
    Else
        sCheck = 1
        gobjRelatorio.sNomeTsk = "pcxnfit"
    End If

    'Verifica o Tipo Destino
    If TipoDestino(1).Value = True Then
        
        lErro = objRelOpcoes.IncluirParametro("NTIPODESTINO", 1)
        If lErro <> AD_BOOL_TRUE Then gError 73667
        
        lErro = objRelOpcoes.IncluirParametro("TDESTINATARIO", FilialEmpresa.Text)
        If lErro <> AD_BOOL_TRUE Then gError 73668
        
        FrameTipo(2).Visible = False
        FrameTipo(1).Visible = True
        
    End If
    
    If TipoDestino(2).Value = True Then
        
        lErro = objRelOpcoes.IncluirParametro("NTIPODESTINO", 2)
        If lErro <> AD_BOOL_TRUE Then gError 73669
    
        lErro = objRelOpcoes.IncluirParametro("NFILIALDESTINO", Codigo_Extrai(FilialFornecedor.Text))
        If lErro <> AD_BOOL_TRUE Then gError 73670
        
        lErro = objRelOpcoes.IncluirParametro("TDESTINATARIO", Fornecedor.Text)
        If lErro <> AD_BOOL_TRUE Then gError 73671
    
        FrameTipo(1).Visible = False
        FrameTipo(2).Visible = True
    End If
    
    If TipoDestino(0).Value = True Then
    
        lErro = objRelOpcoes.IncluirParametro("NTIPODESTINO", 0)
        If lErro <> AD_BOOL_TRUE Then gError 73672
        
        lErro = objRelOpcoes.IncluirParametro("TDESTINATARIO", "0")
        If lErro <> AD_BOOL_TRUE Then gError 73673
        
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NITENS", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 73623

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO

                sOrdenacaoPor = "Codigo"

            Case ORD_POR_DATA

                sOrdenacaoPor = "Data"

            Case ORD_POR_FORNECEDOR

                sOrdenacaoPor = "Fornecedor"
        
            Case ORD_POR_COMPRADOR

                sOrdenacaoPor = "Comprador"

            Case Else
                gError 73624

    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 73625

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 73626

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sNumero_I, sNumero_F, sForn_I, sForn_F, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 73627

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 73609 To 73627, 73659, 73660, 73667, 73668, 73669, 73670, 73671, 73672, 73673

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170744)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sNumero_I As String, sNumero_F As String, sForn_I As String, sForn_F As String, sComp_I As String, sComp_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica Codigo da Filial Inicial e Final
    If CodFilialDe.Text <> "" Then
        sCodFilial_I = CStr(CodFilialDe.Text)
    Else
        sCodFilial_I = ""
    End If

    If CodFilialAte.Text <> "" Then
        sCodFilial_F = CStr(CodFilialAte.Text)
    Else
        sCodFilial_F = ""
    End If

    If sCodFilial_I <> "" And sCodFilial_F <> "" Then

        If StrParaInt(sCodFilial_I) > StrParaInt(sCodFilial_F) Then gError 73628

    End If

    If NomeDe.Text <> "" Then
        sNomeFilial_I = NomeDe.Text
    Else
        sNomeFilial_I = ""
    End If

    If NomeAte.Text <> "" Then
        sNomeFilial_F = NomeAte.Text
    Else
        sNomeFilial_F = ""
    End If

    If sNomeFilial_I <> "" And sNomeFilial_F <> "" Then
        If sNomeFilial_I > sNomeFilial_F Then gError 73629
    End If

    'critica NumeroNF Inicial e Final
    If NumeroDe.Text <> "" Then
        sNumero_I = CStr(NumeroDe.Text)
    Else
        sNumero_I = ""
    End If

    If NumeroAte.Text <> "" Then
        sNumero_F = CStr(NumeroAte.Text)
    Else
        sNumero_F = ""
    End If

    If sNumero_I <> "" And sNumero_F <> "" Then

        If StrParaLong(sNumero_I) > StrParaLong(sNumero_F) Then gError 73630

    End If

    'data inicial não pode ser maior que a final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 73631
    
    End If
    
    'critica Fornecedor Inicial e Final
    If FornecedorDe.Text <> "" Then
        sForn_I = CStr(FornecedorDe.Text)
    Else
        sForn_I = ""
    End If

    If FornecedorAte.Text <> "" Then
        sForn_F = CStr(FornecedorAte.Text)
    Else
        sForn_F = ""
    End If

    If sForn_I <> "" And sForn_F <> "" Then

        If CLng(sForn_I) > CLng(sForn_F) Then gError 73632

    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 73628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodFilialDe.SetFocus

        Case 73629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeDe.SetFocus

        Case 73630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            NumeroDe.SetFocus

        Case 73631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case 73632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornecedorDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170745)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sNumero_I As String, sNumero_F As String, sForn_I As String, sForn_F As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Monta_Expressao_Selecao


   If sCodFilial_I <> "" Then sExpressao = "FilEmpCodInic"

   If sCodFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCodFim"

    End If

   If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeInic"

    End If

    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeFim"

    End If

    If sNumero_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PCCod >= " & Forprint_ConvLong(StrParaLong(sNumero_I))

    End If

    If sNumero_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PCCod <= " & Forprint_ConvLong(StrParaLong(sNumero_F))

    End If
    
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataEnvio >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If

    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataEnvio <= " & Forprint_ConvData(CDate(DataAte.Text))

    End If

    If sForn_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornCod >= " & Forprint_ConvLong(StrParaLong(sForn_I))

    End If

    If sForn_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornCod <= " & Forprint_ConvLong(StrParaLong(sForn_F))

    End If

    'Se a opção para Tipo Destino = FilialEmpresa estiver selecionada
    If TipoDestino(1).Value = True Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoDestino = " & Forprint_ConvInt(StrParaInt(TIPO_DESTINO_EMPRESA))
        sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialDestino = " & Forprint_ConvInt(Codigo_Extrai(FilialEmpresa.Text))
        
    End If
        
    'Se a opção para Tipo Destino = Fornecedor estiver selecionada
    If TipoDestino(2).Value = True Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoDestino = " & Forprint_ConvInt(TIPO_DESTINO_FORNECEDOR)
        sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialDestino = " & Forprint_ConvInt(Codigo_Extrai(FilialFornecedor.Text))
        sExpressao = sExpressao & " E "
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 74996
        
        sExpressao = sExpressao & "FornCliDestino = " & Forprint_ConvLong(objFornecedor.lCodigo)
        
    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case 74996
            'erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170746)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sOrdenacaoPor As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 73634

    'pega Codigo Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73635

    CodFilialDe.Text = sParam
    Call CodFilialDe_Validate(bSGECancelDummy)

    'pega  Codigo Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73636

    CodFilialAte.Text = sParam
    Call CodFilialAte_Validate(bSGECancelDummy)

    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73637

    NomeDe.Text = sParam
    Call NomeDe_Validate(bSGECancelDummy)

    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73638

    NomeAte.Text = sParam
    Call NomeAte_Validate(bSGECancelDummy)

    'pega  Numero inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCINIC", sParam)
    If lErro <> SUCESSO Then gError 73639

    NumeroDe.Text = sParam

    'pega numero final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCFIM", sParam)
    If lErro <> SUCESSO Then gError 73640

    NumeroAte.Text = sParam

    'pega Fornecedor Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 73641

    FornecedorDe.Text = sParam

    'pega Fornecedor Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 73642

    FornecedorAte.Text = sParam

    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAENVINIC", sParam)
    If lErro <> SUCESSO Then gError 73643

    Call DateParaMasked(DataDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAENVFIM", sParam)
    If lErro <> SUCESSO Then gError 73644

    Call DateParaMasked(DataAte, CDate(sParam))

    lErro = objRelOpcoes.ObterParametro("NITENS", sParam)
    If lErro <> SUCESSO Then gError 73645

    If sParam = "1" Then
        CheckItens.Value = 1
    Else
        CheckItens.Value = 0
    End If
    
    'Tipo Destino
    lErro = objRelOpcoes.ObterParametro("NTIPODESTINO", sParam)
    If lErro <> SUCESSO Then gError 73663
    
    If sParam = "0" Then
        TipoDestino(0).Value = vbChecked
    Else
        iIndice = Codigo_Extrai(sParam)
        TipoDestino_Click (iIndice)
        
        If iIndice = 2 Then
        
            lErro = objRelOpcoes.ObterParametro("TDESTINATARIO", sParam)
            If lErro <> SUCESSO Then gError 73664
        
            Fornecedor.Text = sParam
            Call Fornecedor_Validate(bSGECancelDummy)
        
            lErro = objRelOpcoes.ObterParametro("NFILIALDESTINO", sParam)
            If lErro <> SUCESSO Then gError 73665
        
            FilialFornecedor.Text = sParam
            Call FilialFornecedor_Validate(bSGECancelDummy)
            FrameTipo(1).Visible = False
            FrameTipo(2).Visible = True
            TipoDestino(2).Value = vbChecked
            
        ElseIf iIndice = 1 Then
            
            lErro = objRelOpcoes.ObterParametro("TDESTINATARIO", sParam)
            If lErro <> SUCESSO Then gError 73666
            
            FilialEmpresa.Text = sParam
            FrameTipo(2).Visible = False
            FrameTipo(1).Visible = True
            TipoDestino(1).Value = vbChecked
            
        End If
    End If
    
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 73646

    Select Case sOrdenacaoPor

            Case "Comprador"
                
                ComboOrdenacao.ListIndex = ORD_POR_COMPRADOR
                
            Case "Data"

                ComboOrdenacao.ListIndex = ORD_POR_DATA

            Case "Fornecedor"
                
                ComboOrdenacao.ListIndex = ORD_POR_FORNECEDOR

            Case "Codigo"
                
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO

            Case Else
                gError 73647

    End Select

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 73634 To 73647, 73661, 73662, 73663, 73664, 73665, 73666

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170747)

    End Select

    Exit Function

End Function


Private Sub CodFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodFilialDe_Validate

    If Len(Trim(CodFilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 73648

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 73649

    End If

    Exit Sub

Erro_CodFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 73648

        Case 73649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170748)

    End Select

    Exit Sub

End Sub
Private Sub CodFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodFilialAte_Validate

    If Len(Trim(CodFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 73650

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 73651

    End If

    Exit Sub

Erro_CodFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73650

        Case 73651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170749)

    End Select

    Exit Sub

End Sub

Private Sub NomeDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeDe_Validate

    bAchou = False

    If Len(Trim(NomeDe.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 73652

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeDe.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73653

        NomeDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73652

        Case 73653
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170750)

    End Select

Exit Sub

End Sub

Private Sub NomeAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeAte_Validate

    bAchou = False
    If Len(Trim(NomeAte.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 73654

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeAte.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73655

        NomeAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73654

        Case 73655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170751)

    End Select

Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Relação Pedidos de Compra X Notas Fiscais"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpPCNotasFiscais"

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

Public Sub Unload(objme As Object)

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
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is NumeroDe Then
            Call LabelNumeroDe_Click

        ElseIf Me.ActiveControl Is NumeroAte Then
            Call LabelNumeroAte_Click

        ElseIf Me.ActiveControl Is CodFilialDe Then
            Call LabelCodFilialDe_Click

        ElseIf Me.ActiveControl Is CodFilialAte Then
            Call LabelCodFilialAte_Click

        ElseIf Me.ActiveControl Is NomeDe Then
            Call LabelNomeDe_Click

        ElseIf Me.ActiveControl Is NomeAte Then
            Call LabelNomeAte_Click

        ElseIf Me.ActiveControl Is FornecedorDe Then
            Call LabelFornecedorDe_Click

        ElseIf Me.ActiveControl Is FornecedorAte Then
            Call LabelFornecedorAte_Click

        End If

    End If

End Sub



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialEmpDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialEmpDestino, Source, X, Y)
End Sub

Private Sub LabelFilialEmpDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialEmpDestino, Button, Shift, X, Y)
End Sub

Private Sub LabelFornDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornDestino, Source, X, Y)
End Sub

Private Sub LabelFornDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornDestino, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialFornDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialFornDestino, Source, X, Y)
End Sub

Private Sub LabelFilialFornDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialFornDestino, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorAte, Source, X, Y)
End Sub

Private Sub LabelFornecedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorDe, Source, X, Y)
End Sub

Private Sub LabelFornecedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorDe, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelNumeroAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumeroAte, Source, X, Y)
End Sub

Private Sub LabelNumeroAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumeroAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNumeroDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumeroDe, Source, X, Y)
End Sub

Private Sub LabelNumeroDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumeroDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodFilialDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialDe, Source, X, Y)
End Sub

Private Sub LabelCodFilialDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodFilialAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialAte, Source, X, Y)
End Sub

Private Sub LabelCodFilialAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

