VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpConcorrenciasOcx 
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   ScaleHeight     =   5820
   ScaleWidth      =   8655
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4110
      Index           =   2
      Left            =   900
      TabIndex        =   26
      Top             =   1395
      Visible         =   0   'False
      Width           =   6540
      Begin VB.Frame Frame4 
         Caption         =   "Compradores"
         Height          =   2490
         Left            =   270
         TabIndex        =   27
         Top             =   405
         Width           =   5955
         Begin VB.Frame Frame6 
            Caption         =   "Nome Reduzido"
            Height          =   690
            Left            =   225
            TabIndex        =   31
            Top             =   1395
            Width           =   5415
            Begin MSMask.MaskEdBox NomeCompDe 
               Height          =   300
               Left            =   630
               TabIndex        =   4
               Top             =   270
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeCompAte 
               Height          =   300
               Left            =   3345
               TabIndex        =   5
               Top             =   255
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeCompDe 
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
               Left            =   270
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   33
               Top             =   330
               Width           =   315
            End
            Begin VB.Label LabelNomeCompAte 
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
               Left            =   2925
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   32
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Código"
            Height          =   645
            Left            =   225
            TabIndex        =   28
            Top             =   450
            Width           =   5370
            Begin MSMask.MaskEdBox CodCompradorDe 
               Height          =   300
               Left            =   600
               TabIndex        =   2
               Top             =   225
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodCompradorAte 
               Height          =   300
               Left            =   3315
               TabIndex        =   3
               Top             =   225
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodCompradorAte 
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
               Left            =   2910
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   30
               Top             =   285
               Width           =   360
            End
            Begin VB.Label LabelCodCompradorDe 
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
               Left            =   225
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   29
               Top             =   270
               Width           =   315
            End
         End
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpConcorrenciasOcx.ctx":0000
      Left            =   1605
      List            =   "RelOpConcorrenciasOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   2640
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
      Left            =   4500
      Picture         =   "RelOpConcorrenciasOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   135
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6375
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpConcorrenciasOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpConcorrenciasOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpConcorrenciasOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpConcorrenciasOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpConcorrenciasOcx.ctx":0A9A
      Left            =   1620
      List            =   "RelOpConcorrenciasOcx.ctx":0AA7
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4200
      Index           =   1
      Left            =   900
      TabIndex        =   34
      Top             =   1350
      Width           =   6495
      Begin VB.Frame Frame2 
         Caption         =   "Concorrências"
         Height          =   2520
         Left            =   90
         TabIndex        =   36
         Top             =   1575
         Width           =   6270
         Begin VB.Frame Frame12 
            Caption         =   "Descrição"
            Height          =   645
            Left            =   180
            TabIndex        =   51
            Top             =   855
            Width           =   5820
            Begin MSMask.MaskEdBox DescricaoAte 
               Height          =   300
               Left            =   3585
               TabIndex        =   13
               Top             =   225
               Width           =   2085
               _ExtentX        =   3678
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DescricaoDe 
               Height          =   300
               Left            =   570
               TabIndex        =   12
               Top             =   225
               Width           =   2460
               _ExtentX        =   4339
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelDescAte 
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
               Left            =   3195
               TabIndex        =   53
               Top             =   270
               Width           =   360
            End
            Begin VB.Label LabelDescDe 
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
               Left            =   225
               TabIndex        =   52
               Top             =   300
               Width           =   315
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Código"
            Height          =   600
            Left            =   180
            TabIndex        =   48
            Top             =   225
            Width           =   5820
            Begin MSMask.MaskEdBox CodConcorrenciaDe 
               Height          =   300
               Left            =   570
               TabIndex        =   10
               Top             =   210
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodConcorrenciaAte 
               Height          =   300
               Left            =   3585
               TabIndex        =   11
               Top             =   210
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodConcDe 
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
               Left            =   225
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   50
               Top             =   270
               Width           =   315
            End
            Begin VB.Label LabelCodConcAte 
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
               Left            =   3195
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   49
               Top             =   270
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data"
            Height          =   630
            Left            =   180
            TabIndex        =   37
            Top             =   1530
            Width           =   5820
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   315
               Left            =   1755
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   195
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   315
               Left            =   570
               TabIndex        =   14
               Top             =   210
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
               Left            =   4785
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   195
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   315
               Left            =   3600
               TabIndex        =   15
               Top             =   210
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
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
               Left            =   3225
               TabIndex        =   41
               Top             =   270
               Width           =   360
            End
            Begin VB.Label Label4 
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
               Left            =   255
               TabIndex        =   40
               Top             =   270
               Width           =   315
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
            Height          =   315
            Left            =   225
            TabIndex        =   16
            Top             =   2160
            Width           =   2070
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Filial Empresa"
         Height          =   1560
         Left            =   90
         TabIndex        =   35
         Top             =   0
         Width           =   6270
         Begin VB.Frame Frame9 
            Caption         =   "Código"
            Height          =   600
            Left            =   180
            TabIndex        =   45
            Top             =   225
            Width           =   5820
            Begin MSMask.MaskEdBox CodFilialDe 
               Height          =   300
               Left            =   600
               TabIndex        =   6
               Top             =   195
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodFilialAte 
               Height          =   300
               Left            =   3540
               TabIndex        =   7
               Top             =   225
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodFilialAte 
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
               Left            =   3165
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   47
               Top             =   270
               Width           =   360
            End
            Begin VB.Label LabelCodFilialDe 
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
               Left            =   225
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   46
               Top             =   255
               Width           =   315
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Nome"
            Height          =   555
            Left            =   180
            TabIndex        =   42
            Top             =   855
            Width           =   5820
            Begin MSMask.MaskEdBox NomeDe 
               Height          =   300
               Left            =   585
               TabIndex        =   8
               Top             =   180
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeAte 
               Height          =   300
               Left            =   3555
               TabIndex        =   9
               Top             =   180
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeAte 
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
               Left            =   3150
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   44
               Top             =   225
               Width           =   360
            End
            Begin VB.Label LabelNomeDe 
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
               Left            =   225
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   43
               Top             =   225
               Width           =   315
            End
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4560
      Left            =   810
      TabIndex        =   25
      Top             =   1035
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8043
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Concorrência"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comprador"
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
      Left            =   270
      TabIndex        =   24
      Top             =   165
      Width           =   615
   End
   Begin VB.Label Label2 
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
      Left            =   270
      TabIndex        =   23
      Top             =   555
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpConcorrenciasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpConcorrencias
Const ORD_POR_CODIGO = 0
Const ORD_POR_DESCRICAO = 1
Const ORD_POR_DATA = 2

Private WithEvents objEventoCodConcDe As AdmEvento
Attribute objEventoCodConcDe.VB_VarHelpID = -1
Private WithEvents objEventoCodConcAte As AdmEvento
Attribute objEventoCodConcAte.VB_VarHelpID = -1
Private WithEvents objEventoCompradorDe As AdmEvento
Attribute objEventoCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoCompradorAte As AdmEvento
Attribute objEventoCompradorAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeCompradorDe As AdmEvento
Attribute objEventoNomeCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeCompradorAte As AdmEvento
Attribute objEventoNomeCompradorAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 73445

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 73446

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 73445

        Case 73446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167773)

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
    If lErro <> SUCESSO Then gError 73447

    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckItens.Value = vbUnchecked

    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 73447

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167774)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodConcDe = New AdmEvento
    Set objEventoCodConcAte = New AdmEvento
    Set objEventoCompradorDe = New AdmEvento
    Set objEventoCompradorAte = New AdmEvento
    Set objEventoNomeCompradorDe = New AdmEvento
    Set objEventoNomeCompradorAte = New AdmEvento
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento

    iFrameAtual = 1

    ComboOrdenacao.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167775)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCodConcDe = Nothing
    Set objEventoCodConcAte = Nothing
    Set objEventoCompradorDe = Nothing
    Set objEventoCompradorAte = Nothing
    Set objEventoNomeCompradorDe = Nothing
    Set objEventoNomeCompradorAte = Nothing
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing

End Sub

Private Sub CodCompradorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodCompradorAte, iAlterado)
    
End Sub

Private Sub CodCompradorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodCompradorDe, iAlterado)
    
End Sub

Private Sub CodConcorrenciaAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodConcorrenciaAte, iAlterado)
    
End Sub

Private Sub CodConcorrenciaDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodConcorrenciaDe, iAlterado)
    
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

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub LabelCodConcAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_LabelCodConcAte_Click

    If Len(Trim(CodConcorrenciaAte.Text)) > 0 Then
        'Preenche com a Concorrencia da tela
        objConcorrencia.lCodigo = StrParaLong(CodConcorrenciaAte.Text)
    End If

    'Chama Tela ConcorrenciaLista
    Call Chama_Tela("ConcorrenciaLista", colSelecao, objConcorrencia, objEventoCodConcAte)

   Exit Sub

Erro_LabelCodConcAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167776)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodConcDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_LabelCodConcDe_Click

    If Len(Trim(CodConcorrenciaDe.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objConcorrencia.lCodigo = StrParaLong(CodConcorrenciaDe.Text)
    End If

    'Chama Tela ConcorrenciaLista
    Call Chama_Tela("ConcorrenciaLista", colSelecao, objConcorrencia, objEventoCodConcDe)

   Exit Sub

Erro_LabelCodConcDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167777)

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
    If lErro <> SUCESSO Then gError 73448

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73448
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167778)

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
    If lErro <> SUCESSO Then gError 73449

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73449
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167779)

    End Select

    Exit Sub

End Sub


Private Sub NomeCompDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuario As New ClassUsuario
Dim objComprador As New ClassComprador

On Error GoTo Erro_NomeCompDe_Validate

    If Len(Trim(NomeCompDe.Text)) > 0 Then
    
        objUsuario.sNomeReduzido = NomeCompDe.Text
    
        'LÊ o usuário
         lErro = CF("Usuario_Le_NomeRed", objUsuario)
         If lErro <> SUCESSO And lErro <> 57269 Then gError 73450
         If lErro = 57269 Then gError 73451
    
        objComprador.sCodUsuario = objUsuario.sCodUsuario
        
        'Lê o Comprador
        lErro = CF("Comprador_Le_Usuario", objComprador)
        If lErro <> SUCESSO And lErro <> 50059 Then gError 73452
        If lErro <> SUCESSO Then gError 73453
    
        NomeCompDe.Text = objUsuario.sNomeReduzido
        
    End If
    
    Exit Sub
    
Erro_NomeCompDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 73450, 73452
            'Erros tratados nas rotinas chamadas
            
        Case 73451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", gErr, objUsuario.sNomeReduzido)
            
        Case 73453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR2", gErr, objUsuario.sCodUsuario)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167780)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub NomeCompAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuario As New ClassUsuario
Dim objComprador As New ClassComprador

On Error GoTo Erro_NomeCompAte_Validate

    If Len(Trim(NomeCompAte.Text)) > 0 Then
    
        objUsuario.sNomeReduzido = NomeCompAte.Text
    
        'LÊ o usuário
         lErro = CF("Usuario_Le_NomeRed", objUsuario)
         If lErro <> SUCESSO And lErro <> 57269 Then gError 73454
         If lErro = 57269 Then gError 73455
    
        objComprador.sCodUsuario = objUsuario.sCodUsuario
        
        'Lê o Comprador
        lErro = CF("Comprador_Le_Usuario", objComprador)
        If lErro <> SUCESSO And lErro <> 50059 Then gError 73456
        If lErro <> SUCESSO Then gError 73457
    
        NomeCompAte.Text = objUsuario.sNomeReduzido

    End If
    
    Exit Sub
    
Erro_NomeCompAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 73454, 73456
            'Erros tratados nas rotinas chamadas
            
        Case 73455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", gErr, objUsuario.sNomeReduzido)
            
        Case 73457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR2", gErr, objUsuario.sCodUsuario)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167781)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoNomeCompradorAte_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    NomeCompAte.Text = objComprador.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeCompradorDe_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    NomeCompDe.Text = objComprador.sNomeReduzido

    Me.Show

    Exit Sub

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

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73458

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 73458
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 167782)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73459

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 73459
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 167783)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73460

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 73460
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 167784)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73461

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 73461
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 167785)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167786)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167787)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodCompradorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCodCompradorAte_Click

    If Len(Trim(CodCompradorAte.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.iCodigo = StrParaInt(CodCompradorAte.Text)
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCompradorAte)

   Exit Sub

Erro_LabelCodCompradorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167788)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodCompradorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCodCompradorDe_Click

    If Len(Trim(CodCompradorDe.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.iCodigo = StrParaInt(CodCompradorDe.Text)
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCompradorDe)

   Exit Sub

Erro_LabelCodCompradorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167789)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeCompDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelNomeCompDe_Click

    If Len(Trim(NomeCompDe.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.sNome = NomeCompDe.Text
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoNomeCompradorDe)

   Exit Sub

Erro_LabelNomeCompDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167790)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeCompAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelNomeCompAte_Click

    If Len(Trim(NomeCompAte.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.sNome = NomeCompAte.Text
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoNomeCompradorAte)

   Exit Sub

Erro_LabelNomeCompAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167791)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167792)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167793)

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

Private Sub objEventoCodConcAte_evSelecao(obj1 As Object)

Dim objConcorrencia As New ClassConcorrencia

    Set objConcorrencia = obj1

    CodConcorrenciaAte.Text = CStr(objConcorrencia.lCodigo)

    Me.Show

End Sub

Private Sub objEventoCodConcDe_evSelecao(obj1 As Object)

Dim objConcorrencia As New ClassConcorrencia

    Set objConcorrencia = obj1

    CodConcorrenciaDe.Text = CStr(objConcorrencia.lCodigo)


    Me.Show

End Sub

Private Sub objEventoCompradorDe_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CodCompradorDe.Text = CStr(objComprador.iCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCompradorAte_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CodCompradorAte.Text = CStr(objComprador.iCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 73462

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73463

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 73464

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 73465

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 73462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 73463 To 73465

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167794)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 73466

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 73467

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 73466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 73467

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167795)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73468

Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "ConcorrenciaCod", 1)
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemConcorrencia", 1)
            
            Case ORD_POR_DESCRICAO

                'Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "Descricao", 1)
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "ConcorrenciaCod", 1)
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemConcorrencia", 1)
            
            Case ORD_POR_DATA
                
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "DataConc", 1)
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "ConcorrenciaCod", 1)
                'Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemConcorrencia", 1)
                
            Case Else
                gError 74946

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 72671, 74946

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167796)

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
Dim sDesc_I As String
Dim sDesc_F As String
Dim sCodConc_I As String
Dim sCodConc_F As String
Dim sNomeComp_I As String
Dim sNomeComp_F As String
Dim sCodComprador_I As String
Dim sCodComprador_F As String
Dim sCheck As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodConc_I, sCodConc_F, sDesc_I, sDesc_F, sCodComprador_I, sCodComprador_F, sNomeComp_I, sNomeComp_F)
    If lErro <> SUCESSO Then gError 73469

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 73470

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 73471

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73472

    lErro = objRelOpcoes.IncluirParametro("NCODCONCINIC", sCodConc_I)
    If lErro <> AD_BOOL_TRUE Then gError 73473

    lErro = objRelOpcoes.IncluirParametro("TDESCCONCINIC", DescricaoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73474

    lErro = objRelOpcoes.IncluirParametro("NCODCOMPINIC", sCodComprador_I)
    If lErro <> AD_BOOL_TRUE Then gError 73475

    lErro = objRelOpcoes.IncluirParametro("TNOMECOMPINIC", sNomeComp_I)
    If lErro <> AD_BOOL_TRUE Then gError 73476

    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATACONCINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATACONCINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73477

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 73478

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73479

    lErro = objRelOpcoes.IncluirParametro("NCODCONCFIM", sCodConc_F)
    If lErro <> AD_BOOL_TRUE Then gError 73480

    lErro = objRelOpcoes.IncluirParametro("TDESCCONCFIM", DescricaoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73481

    lErro = objRelOpcoes.IncluirParametro("NCODCOMPFIM", sCodComprador_F)
    If lErro <> AD_BOOL_TRUE Then gError 73482

    lErro = objRelOpcoes.IncluirParametro("TNOMECOMPFIM", sNomeComp_F)
    If lErro <> AD_BOOL_TRUE Then gError 73483

    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATACONCFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATACONCFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73484

    'Exibe Itens
    If CheckItens.Value = 0 Then
        sCheck = 0
        gobjRelatorio.sNomeTsk = "cotrec"
    Else
        sCheck = 1
        gobjRelatorio.sNomeTsk = "cotrecit"
    End If

    lErro = objRelOpcoes.IncluirParametro("NITENS", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 73485

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO

                sOrdenacaoPor = "Codigo"

            Case ORD_POR_DESCRICAO

                sOrdenacaoPor = "Descricao"

            Case ORD_POR_DATA
                sOrdenacaoPor = "Data"

            Case Else
                gError 73486

    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 73487

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 73488

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodConc_I, sCodConc_F, sDesc_I, sDesc_F, sCodComprador_I, sCodComprador_F, sNomeComp_I, sNomeComp_F, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 73489

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 73469 To 73489

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167797)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodConc_I As String, sCodConc_F As String, sDesc_I As String, sDesc_F As String, sCodComprador_I As String, sCodComprador_F As String, sNomeComprador_I As String, sNomeComprador_F As String) As Long
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

        If StrParaInt(sCodFilial_I) > StrParaInt(sCodFilial_F) Then gError 73490

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
        If sNomeFilial_I > sNomeFilial_F Then gError 73491
    End If

    'critica CodigoConc Inicial e Final
    If CodConcorrenciaDe.Text <> "" Then
        sCodConc_I = CStr(CodConcorrenciaDe.Text)
    Else
        sCodConc_I = ""
    End If

    If CodConcorrenciaAte.Text <> "" Then
        sCodConc_F = CStr(CodConcorrenciaAte.Text)
    Else
        sCodConc_F = ""
    End If

    If sCodConc_I <> "" And sCodConc_F <> "" Then

        If StrParaLong(sCodConc_I) > StrParaLong(sCodConc_F) Then gError 73492

    End If

    'data inicial não pode ser maior que a final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 73493
    
    End If
    
    'critica Comprador Inicial e Final
    If CodCompradorDe.Text <> "" Then
        sCodComprador_I = CStr(CodCompradorDe.Text)
    Else
        sCodComprador_I = ""
    End If

    If CodCompradorAte.Text <> "" Then
        sCodComprador_F = CStr(CodCompradorAte.Text)
    Else
        sCodComprador_F = ""
    End If

    If sCodComprador_I <> "" And sCodComprador_F <> "" Then

        If StrParaInt(sCodComprador_I) > StrParaInt(sCodComprador_F) Then gError 73494

    End If

    'critica Comprador Inicial e Final
    If NomeCompDe.Text <> "" Then
        sNomeComprador_I = CStr(NomeCompDe.Text)
    Else
        sNomeComprador_I = ""
    End If

    If NomeCompAte.Text <> "" Then
        sNomeComprador_F = CStr(NomeCompAte.Text)
    Else
        sNomeComprador_F = ""
    End If

    If sNomeComprador_I <> "" And sNomeComprador_F <> "" Then

        If sNomeComprador_I > sNomeComprador_F Then gError 73495

    End If
    'critica Descricao Inicial e Final
    If DescricaoDe.Text <> "" Then
        sDesc_I = CStr(DescricaoDe.Text)
    Else
        sDesc_I = ""
    End If

    If DescricaoAte.Text <> "" Then
        sDesc_F = CStr(DescricaoAte.Text)
    Else
        sDesc_F = ""
    End If

    If sDesc_I <> "" And sDesc_F <> "" Then

        If sDesc_I > sDesc_F Then gError 73496

    End If


    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 73490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodFilialDe.SetFocus

        Case 73491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeDe.SetFocus

        Case 73492
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            CodConcorrenciaDe.SetFocus

        Case 73493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
            
        Case 73494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            CodCompradorDe.SetFocus

        Case 73495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            NomeCompDe.SetFocus

        Case 73496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_INICIAL_MAIOR", gErr)
            DescricaoDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167798)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodConc_I As String, sCodConc_F As String, sDesc_I As String, sDesc_F As String, sCodComprador_I As String, sCodComprador_F As String, sNomeComprador_I As String, sNomeComprador_F As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao


   If sCodFilial_I <> "" Then sExpressao = "FilialEmpresaCod >= " & Forprint_ConvInt(StrParaInt(sCodFilial_I))

   If sCodFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaCod <= " & Forprint_ConvInt(StrParaInt(sCodFilial_F))

    End If

   If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaNome >= " & Forprint_ConvTexto(sNomeFilial_I)

    End If

    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaNome <= " & Forprint_ConvTexto(sNomeFilial_F)

    End If

    If sCodConc_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ConcorrenciaCod >= " & Forprint_ConvLong(StrParaLong(sCodConc_I))

    End If

    If sCodConc_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ConcorrenciaCod <= " & Forprint_ConvLong(StrParaLong(sCodConc_F))

    End If
    
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataConc >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If

    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataConc <= " & Forprint_ConvData(CDate(DataAte.Text))

    End If

    If sNomeComprador_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "UsuarioNome >= " & Forprint_ConvTexto((sNomeComprador_I))

    End If

    If sNomeComprador_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "UsuarioNome <= " & Forprint_ConvTexto((sNomeComprador_F))

    End If

    If sDesc_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Descricao >= " & Forprint_ConvTexto((sDesc_I))

    End If

    If sDesc_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Descricao <= " & Forprint_ConvTexto((sDesc_F))

    End If

    If sCodComprador_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CompradorCod >= " & Forprint_ConvInt(StrParaInt(sCodComprador_I))

    End If

    If sCodComprador_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CompradorCod <= " & Forprint_ConvInt(StrParaInt(sCodComprador_F))

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167799)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sOrdenacaoPor As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 73497

    'pega Codigo Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73498

    CodFilialDe.Text = sParam
    Call CodFilialDe_Validate(bSGECancelDummy)

    'pega  Codigo Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73499

    CodFilialAte.Text = sParam
    Call CodFilialAte_Validate(bSGECancelDummy)

    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73500

    NomeDe.Text = sParam
    Call NomeDe_Validate(bSGECancelDummy)

    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73501

    NomeAte.Text = sParam
    Call NomeAte_Validate(bSGECancelDummy)

    'pega  Codigo Conc inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCONCINIC", sParam)
    If lErro <> SUCESSO Then gError 73502

    CodConcorrenciaDe.Text = sParam

    'pega  Codigo Conc final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCONCFIM", sParam)
    If lErro <> SUCESSO Then gError 73503

    CodConcorrenciaAte.Text = sParam

    'pega Comprador Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCOMPINIC", sParam)
    If lErro <> SUCESSO Then gError 73504

    CodCompradorDe.Text = sParam

    'pega Comprador Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCOMPFIM", sParam)
    If lErro <> SUCESSO Then gError 73505

    CodCompradorAte.Text = sParam

    'pega  Nome do Comprador Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMECOMPINIC", sParam)
    If lErro <> SUCESSO Then gError 73506

    NomeCompDe.Text = sParam
    Call NomeCompDe_Validate(bSGECancelDummy)

    'pega nome do comprador Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMECOMPFIM", sParam)
    If lErro <> SUCESSO Then gError 73507

    NomeCompAte.Text = sParam
    Call NomeCompAte_Validate(bSGECancelDummy)

    'pega Descricao Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TDESCCONCINIC", sParam)
    If lErro <> SUCESSO Then gError 73508

    DescricaoDe.Text = sParam

    'pega Descricao Final e exibe
    lErro = objRelOpcoes.ObterParametro("TDESCCONCFIM", sParam)
    If lErro <> SUCESSO Then gError 73509

    DescricaoAte.Text = sParam
    
    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATACONCINIC", sParam)
    If lErro <> SUCESSO Then gError 73510

    Call DateParaMasked(DataDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATACONCFIM", sParam)
    If lErro <> SUCESSO Then gError 73511

    Call DateParaMasked(DataAte, CDate(sParam))

    lErro = objRelOpcoes.ObterParametro("NITENS", sParam)
    If lErro <> SUCESSO Then gError 73512

    If sParam = "1" Then
        CheckItens.Value = 1
    Else
        CheckItens.Value = 0
    End If
    
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 73513

    Select Case sOrdenacaoPor

            Case "Codigo"

                ComboOrdenacao.ListIndex = ORD_POR_CODIGO

            Case "Descricao"
                
                ComboOrdenacao.ListIndex = ORD_POR_DESCRICAO

            Case "Data"
                
                ComboOrdenacao.ListIndex = ORD_POR_DATA

            Case Else
                gError 73514

    End Select

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 73497 To 73514

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167800)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 73515

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 73516

    End If

    Exit Sub

Erro_CodFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 73515

        Case 73516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167801)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 73517

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 73518

    End If

    Exit Sub

Erro_CodFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73517

        Case 73518
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167802)

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
        If lErro <> SUCESSO Then gError 73519

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeDe.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73520

        NomeDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73519

        Case 73520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167803)

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
        If lErro <> SUCESSO Then gError 73521

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeAte.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73522

        NomeAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73521

        Case 73522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167804)

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
    Caption = "Análise de Cotações Recebidas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpConcorrencias"

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

        If Me.ActiveControl Is CodConcorrenciaDe Then
            Call LabelCodConcDe_Click

        ElseIf Me.ActiveControl Is CodConcorrenciaAte Then
            Call LabelCodConcAte_Click

        ElseIf Me.ActiveControl Is CodFilialDe Then
            Call LabelCodFilialDe_Click

        ElseIf Me.ActiveControl Is CodFilialAte Then
            Call LabelCodFilialAte_Click

        ElseIf Me.ActiveControl Is NomeDe Then
            Call LabelNomeDe_Click

        ElseIf Me.ActiveControl Is NomeAte Then
            Call LabelNomeAte_Click

        ElseIf Me.ActiveControl Is CodCompradorDe Then
            Call LabelCodCompradorDe_Click

        ElseIf Me.ActiveControl Is CodCompradorAte Then
            Call LabelCodCompradorAte_Click

        ElseIf Me.ActiveControl Is NomeCompDe Then
            Call LabelNomeCompDe_Click

        ElseIf Me.ActiveControl Is NomeCompAte Then
            Call LabelNomeCompAte_Click

        End If

    End If

End Sub


Private Sub LabelCodFilialDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialDe, Source, X, Y)
End Sub

Private Sub LabelCodFilialDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodFilialAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialAte, Source, X, Y)
End Sub

Private Sub LabelCodFilialAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialAte, Button, Shift, X, Y)
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

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodCompradorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodCompradorAte, Source, X, Y)
End Sub

Private Sub LabelCodCompradorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodCompradorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodCompradorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodCompradorDe, Source, X, Y)
End Sub

Private Sub LabelCodCompradorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodCompradorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeCompDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeCompDe, Source, X, Y)
End Sub

Private Sub LabelNomeCompDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeCompDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeCompAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeCompAte, Source, X, Y)
End Sub

Private Sub LabelNomeCompAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeCompAte, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelCodConcAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodConcAte, Source, X, Y)
End Sub

Private Sub LabelCodConcAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodConcAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodConcDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodConcDe, Source, X, Y)
End Sub

Private Sub LabelCodConcDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodConcDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDescAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescAte, Source, X, Y)
End Sub

Private Sub LabelDescAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDescDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescDe, Source, X, Y)
End Sub

Private Sub LabelDescDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescDe, Button, Shift, X, Y)
End Sub

