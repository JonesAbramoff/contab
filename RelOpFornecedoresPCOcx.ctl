VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpFornecedoresPCOcx 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   ScaleHeight     =   5760
   ScaleWidth      =   8430
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4200
      Index           =   2
      Left            =   720
      TabIndex        =   25
      Top             =   1305
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Frame Frame2 
         Caption         =   "Pedidos de Compra"
         Height          =   2835
         Left            =   180
         TabIndex        =   31
         Top             =   1215
         Width           =   6120
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
            Left            =   210
            TabIndex        =   57
            Top             =   2460
            Width           =   2685
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data"
            Height          =   630
            Left            =   225
            TabIndex        =   40
            Top             =   990
            Width           =   4470
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   315
               Left            =   1665
               TabIndex        =   41
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
               Left            =   480
               TabIndex        =   8
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
               Left            =   3975
               TabIndex        =   42
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
               Left            =   2790
               TabIndex        =   9
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
               Left            =   2415
               TabIndex        =   44
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
               Left            =   165
               TabIndex        =   43
               Top             =   270
               Width           =   315
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Código"
            Height          =   615
            Left            =   225
            TabIndex        =   37
            Top             =   240
            Width           =   4455
            Begin MSMask.MaskEdBox CodPCDe 
               Height          =   300
               Left            =   525
               TabIndex        =   6
               Top             =   195
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodPCAte 
               Height          =   300
               Left            =   2790
               TabIndex        =   7
               Top             =   195
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodPCAte 
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
               Left            =   2415
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   39
               Top             =   255
               Width           =   360
            End
            Begin VB.Label LabelCodPCDe 
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
               TabIndex        =   38
               Top             =   255
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Data de Envio"
            Height          =   645
            Left            =   225
            TabIndex        =   32
            Top             =   1755
            Width           =   4500
            Begin MSComCtl2.UpDown UpDownDataEnvioDe 
               Height          =   315
               Left            =   1680
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   225
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEnvioDe 
               Height          =   315
               Left            =   495
               TabIndex        =   10
               Top             =   240
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataEnvioAte 
               Height          =   315
               Left            =   3960
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   210
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEnvioAte 
               Height          =   315
               Left            =   2790
               TabIndex        =   11
               Top             =   225
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeReqDe 
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
            Begin VB.Label LabelNomeReqAte 
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
               Left            =   2415
               TabIndex        =   35
               Top             =   315
               Width           =   360
            End
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Filial Empresa"
         Height          =   1095
         Left            =   180
         TabIndex        =   26
         Top             =   45
         Width           =   6135
         Begin MSMask.MaskEdBox CodFilialDe 
            Height          =   300
            Left            =   1080
            TabIndex        =   2
            Top             =   270
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
            Left            =   4095
            TabIndex        =   3
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeDe 
            Height          =   300
            Left            =   1080
            TabIndex        =   4
            Top             =   675
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeAte 
            Height          =   300
            Left            =   4080
            TabIndex        =   5
            Top             =   660
            Width           =   1890
            _ExtentX        =   3334
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
            Left            =   3165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   705
            Width           =   900
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
            Left            =   195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   29
            Top             =   705
            Width           =   855
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
            Left            =   3015
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   315
            Width           =   1005
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
            Left            =   105
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   315
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4200
      Index           =   1
      Left            =   765
      TabIndex        =   45
      Top             =   1260
      Width           =   6495
      Begin VB.Frame Frame4 
         Caption         =   "Fornecedores"
         Height          =   3810
         Left            =   210
         TabIndex        =   46
         Top             =   210
         Width           =   5895
         Begin VB.Frame Frame9 
            Caption         =   "Tipo"
            Height          =   1215
            Left            =   300
            TabIndex        =   53
            Top             =   2370
            Width           =   5115
            Begin VB.ComboBox ComboTipo 
               Height          =   315
               Left            =   1800
               TabIndex        =   56
               Top             =   570
               Width           =   3195
            End
            Begin VB.OptionButton OptionUmTipo 
               Caption         =   "Apenas do Tipo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   60
               TabIndex        =   55
               Top             =   645
               Width           =   1755
            End
            Begin VB.OptionButton OptionTodosTipos 
               Caption         =   "Todos os  Tipos"
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
               Left            =   75
               TabIndex        =   54
               Top             =   330
               Width           =   1890
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Código"
            Height          =   690
            Left            =   300
            TabIndex        =   50
            Top             =   405
            Width           =   5085
            Begin MSMask.MaskEdBox CodFornecedorDe 
               Height          =   300
               Left            =   570
               TabIndex        =   12
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodFornecedorAte 
               Height          =   300
               Left            =   2925
               TabIndex        =   13
               Top             =   240
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodFornecedorAte 
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
               Left            =   2520
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   52
               Top             =   315
               Width           =   360
            End
            Begin VB.Label LabelCodFornecedorDe 
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
               Left            =   240
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   51
               Top             =   315
               Width           =   315
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Nome Reduzido"
            Height          =   720
            Left            =   300
            TabIndex        =   47
            Top             =   1440
            Width           =   5085
            Begin MSMask.MaskEdBox NomeFornDe 
               Height          =   300
               Left            =   555
               TabIndex        =   14
               Top             =   255
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeFornAte 
               Height          =   300
               Left            =   2940
               TabIndex        =   15
               Top             =   255
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeFornDe 
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
               TabIndex        =   49
               Top             =   315
               Width           =   315
            End
            Begin VB.Label LabelNomeFornAte 
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
               Left            =   2520
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   48
               Top             =   315
               Width           =   360
            End
         End
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpFornecedoresPCOcx.ctx":0000
      Left            =   1575
      List            =   "RelOpFornecedoresPCOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   2460
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
      Left            =   4275
      Picture         =   "RelOpFornecedoresPCOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   135
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6165
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpFornecedoresPCOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpFornecedoresPCOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpFornecedoresPCOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpFornecedoresPCOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpFornecedoresPCOcx.ctx":0A9A
      Left            =   1575
      List            =   "RelOpFornecedoresPCOcx.ctx":0AAA
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Visible         =   0   'False
      Width           =   2460
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4605
      Left            =   705
      TabIndex        =   24
      Top             =   945
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8123
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fornecedor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedido"
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
      Left            =   210
      TabIndex        =   23
      Top             =   150
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
      Left            =   210
      TabIndex        =   22
      Top             =   540
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpFornecedoresPCOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***** Alteração feita em 18/04/01 ***********
'A combo Ordenacao não está visível em tempo de execução, pois o relatório não está preparado para ordenar
'Para tornar essa combo visível, será necessário verificar se o código para ordenação está correto,
'e alterar o relatório, deixando-o preparado para aceitar as possíveis ordenações
'***** Feito por Luiz Gustavo *****************

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpFornecedoresPC
Const ORD_POR_FORNECEDOR = 0
Const ORD_POR_NOME_FORNECEDOR = 1
Const ORD_POR_DATAENVIO = 2
Const ORD_POR_COMPRADOR = 3


Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorDe As AdmEvento
Attribute objEventoFornecedorDe.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorAte As AdmEvento
Attribute objEventoFornecedorAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFornecedorDe As AdmEvento
Attribute objEventoNomeFornecedorDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFornecedorAte As AdmEvento
Attribute objEventoNomeFornecedorAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1


Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 72571

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 72570

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 72570

        Case 72571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169264)

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
    If lErro <> SUCESSO Then gError 72572

    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckItens.Value = vbUnchecked
    
    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 72572

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169265)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodPCDe = New AdmEvento
    Set objEventoCodPCAte = New AdmEvento
    Set objEventoFornecedorDe = New AdmEvento
    Set objEventoFornecedorAte = New AdmEvento
    Set objEventoNomeFornecedorDe = New AdmEvento
    Set objEventoNomeFornecedorAte = New AdmEvento
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
    
    'Seleciona a opção Todos os tipos de Fornecedor
    OptionTodosTipos.Value = True
    
    'Desabilita a Combo para seleção do tipo de Fornecedor
    ComboTipo.Enabled = False
    
    'carrega combo de tipo fornecedor
    lErro = CF("TipoFornecedor_CarregaCombo", ComboTipo)
    If lErro <> SUCESSO Then gError 72731
        
    iFrameAtual = 1
        
    ComboOrdenacao.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 72731
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169266)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
    Set objEventoFornecedorDe = Nothing
    Set objEventoFornecedorAte = Nothing
    Set objEventoNomeFornecedorDe = Nothing
    Set objEventoNomeFornecedorAte = Nothing
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing

End Sub

Private Sub CodFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialAte, iAlterado)
    
End Sub

Private Sub CodFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialDe, iAlterado)
    
End Sub

Private Sub CodFornecedorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFornecedorAte, iAlterado)
    
End Sub

Private Sub CodFornecedorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFornecedorDe, iAlterado)
    
End Sub

Private Sub CodPCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCAte, iAlterado)
    
End Sub

Private Sub CodPCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCDe, iAlterado)
    
End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    
End Sub

Private Sub DataEnvioAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioAte, iAlterado)
    
End Sub

Private Sub DataEnvioDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioDe, iAlterado)
    
End Sub

Private Sub LabelCodPCAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelCodPCAte_Click

    If Len(Trim(CodPCAte.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedCompra.lCodigo = StrParaLong(CodPCAte.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoCodPCAte)

   Exit Sub

Erro_LabelCodPCAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169267)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodPCDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelCodPCDe_Click

    If Len(Trim(CodPCDe.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedCompra.lCodigo = StrParaLong(CodPCDe.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoCodPCDe)

   Exit Sub

Erro_LabelCodPCDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169268)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEnvioDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEnvioDe.Text)
    If lErro <> SUCESSO Then gError 72573

    Exit Sub

Erro_DataEnvioDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72573
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169269)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEnvioAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEnvioAte.Text)
    If lErro <> SUCESSO Then gError 72574

    Exit Sub

Erro_DataEnvioAte_Validate:

    Cancel = True

    Select Case gErr

        Case 72574
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169270)

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
    If lErro <> SUCESSO Then gError 72575

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 72575
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169271)

    End Select

    Exit Sub

End Sub

Private Sub NomeFornDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NomeFornDe_Validate

    If Len(Trim(NomeFornDe.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = NomeFornDe.Text
    
        'LÊ o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 72732
        
        If lErro <> SUCESSO Then gError 72733
        
    End If
    
    Exit Sub
    
Erro_NomeFornDe_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 72732
        
        Case 72733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169272)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub NomeFornAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NomeFornAte_Validate

    If Len(Trim(NomeFornAte.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = NomeFornAte.Text
    
        'LÊ o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 72734
        If lErro <> SUCESSO Then gError 72735
        
    End If
    
    Exit Sub
    
Erro_NomeFornAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 72734
        
        Case 72735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169273)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoNomeFornecedorAte_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    NomeFornAte.Text = objFornecedor.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFornecedorDe_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    NomeFornDe.Text = objFornecedor.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub OptionUmTipo_Click()
    ComboTipo.Enabled = True
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


Private Sub UpDownDataEnvioDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 72576

    Exit Sub

Erro_UpDownDataEnvioDe_DownClick:

    Select Case gErr

        Case 72576
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169274)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 72577

    Exit Sub

Erro_UpDownDataEnvioDe_UpClick:

    Select Case gErr

        Case 72577
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169275)

    End Select

    Exit Sub

End Sub
Private Sub UpDownDataEnvioAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 72578

    Exit Sub

Erro_UpDownDataEnvioAte_DownClick:

    Select Case gErr

        Case 72578
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169276)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 72579

    Exit Sub

Erro_UpDownDataEnvioAte_UpClick:

    Select Case gErr

        Case 72579
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169277)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 72580

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 72580
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169278)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 72581

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 72581
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169279)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 72582

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 72582
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169280)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 72583

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 72583
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169281)

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
    If lErro <> SUCESSO Then gError 72584

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72584
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169282)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169283)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169284)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodFornecedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelCodFornecedorAte_Click

    If Len(Trim(CodFornecedorAte.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(CodFornecedorAte.Text)
    End If

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorAte)

   Exit Sub

Erro_LabelCodFornecedorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169285)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodFornecedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelCodFornecedorDe_Click

    If Len(Trim(CodFornecedorDe.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(CodFornecedorDe.Text)
    End If

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorDe)

   Exit Sub

Erro_LabelCodFornecedorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169286)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeFornDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNomeFornDe_Click

    If Len(Trim(NomeFornDe.Text)) > 0 Then
        'Preenche com o comprador da tela
        objFornecedor.sNomeReduzido = NomeFornDe.Text
    End If

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoNomeFornecedorDe)

   Exit Sub

Erro_LabelNomeFornDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169287)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeFornAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNomeFornAte_Click

    If Len(Trim(NomeFornAte.Text)) > 0 Then
        'Preenche com o comprador da tela
        objFornecedor.sNomeReduzido = NomeFornAte.Text
    End If

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoNomeFornecedorAte)

   Exit Sub

Erro_LabelNomeFornAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169288)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169289)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169290)

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

Private Sub objEventoCodPCAte_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    CodPCAte.Text = CStr(objPedCompra.lCodigo)

    Me.Show

End Sub

Private Sub objEventoCodPCDe_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    CodPCDe.Text = CStr(objPedCompra.lCodigo)

    Me.Show

End Sub

Private Sub objEventoFornecedorDe_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    CodFornecedorDe.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornecedorAte_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    CodFornecedorAte.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 72585

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 72586

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 72587

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 72588

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 72585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 72586 To 72588

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169291)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 72589

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 72590

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 72589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 72590

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169292)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 72591

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_FORNECEDOR
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)

'            Case ORD_POR_NOME_FORNECEDOR
'
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaNome", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorNome", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornNome", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)
'
'            Case ORD_POR_DATAENVIO
'
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEnvio", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)
'
'            Case ORD_POR_COMPRADOR
'
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCComprador", 1)
'                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCompraCod", 1)
'
            Case Else
                gError 74947

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 72591, 74497

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169293)

    End Select

    Exit Sub

End Sub

Private Sub OptionTodosTipos_Click()

Dim lErro As Long

On Error GoTo Erro_OptionTodosTipos_Click

    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = False
    'OptionTodosTipos.Value = True
    
    Exit Sub

Erro_OptionTodosTipos_Click:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169294)

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
Dim sCodPC_I As String
Dim sCodPC_F As String
Dim sNomeForn_I As String
Dim sNomeForn_F As String
Dim sCodFornecedor_I As String
Dim sCodFornecedor_F As String
Dim sCheck As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String
Dim sCheckTipo As String
Dim sFornecedorTipo As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodPC_I, sCodPC_F, sCodFornecedor_I, sCodFornecedor_F, sNomeForn_I, sNomeForn_F)
    If lErro <> SUCESSO Then gError 72592

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 72593

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 72594

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72595

    lErro = objRelOpcoes.IncluirParametro("NCODPCINIC", sCodPC_I)
    If lErro <> AD_BOOL_TRUE Then gError 72596

    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sCodFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then gError 72597

    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNINIC", sNomeForn_I)
    If lErro <> AD_BOOL_TRUE Then gError 72598

    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72599

    'Preenche a data envio inicial
    If Trim(DataEnvioDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", DataEnvioDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72600

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 72601

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72602

    lErro = objRelOpcoes.IncluirParametro("NCODPCFIM", sCodPC_F)
    If lErro <> AD_BOOL_TRUE Then gError 72603

    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sCodFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then gError 72604

    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNFIM", sNomeForn_F)
    If lErro <> AD_BOOL_TRUE Then gError 72605

    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72606

    'Preenche a data envio final
    If Trim(DataEnvioAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", DataEnvioAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72607

    'Exibe Itens
    If CheckItens.Value = 0 Then
        sCheck = 0
        gobjRelatorio.sNomeTsk = "fornpc"
    Else
        sCheck = 1
        gobjRelatorio.sNomeTsk = "fornpcit"
    End If

    lErro = objRelOpcoes.IncluirParametro("NITENS", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 72608

    Select Case ComboOrdenacao.ListIndex
            
            Case ORD_POR_COMPRADOR

                sOrdenacaoPor = "CodComprador"

            Case ORD_POR_NOME_FORNECEDOR

                sOrdenacaoPor = "NomeFornecedor"

            Case ORD_POR_DATAENVIO
                sOrdenacaoPor = "DataEnvio"

            Case ORD_POR_FORNECEDOR

                sOrdenacaoPor = "Fornecedor"


            Case Else
                gError 72609

    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 72610

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 72611

    lErro = objRelOpcoes.IncluirParametro("NTIPOFORN", Codigo_Extrai(ComboTipo.Text))
    If lErro <> AD_BOOL_TRUE Then gError 79992
    
    lErro = objRelOpcoes.IncluirParametro("TTFORNECEDOR", ComboTipo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72727

    If OptionTodosTipos.Value = True Then
        sCheckTipo = "Todos"
    
    ElseIf OptionUmTipo = True Then
    
        sCheckTipo = "Um"
    
    Else
    
        gError 79991
    
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 72728

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodPC_I, sCodPC_F, sCodFornecedor_I, sCodFornecedor_F, sNomeForn_I, sNomeForn_F, sFornecedorTipo, sCheckTipo, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 72612

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 72592 To 72612, 72726, 72727, 72728, 79992

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169295)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodPC_I As String, sCodPC_F As String, sCodFornecedor_I As String, sCodFornecedor_F As String, sNomeFornecedor_I As String, sNomeFornecedor_F As String) As Long
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

        If StrParaInt(sCodFilial_I) > StrParaInt(sCodFilial_F) Then gError 72613

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
        If sNomeFilial_I > sNomeFilial_F Then gError 72614
    End If

    'critica CodigoPC Inicial e Final
    If CodPCDe.Text <> "" Then
        sCodPC_I = CStr(CodPCDe.Text)
    Else
        sCodPC_I = ""
    End If

    If CodPCAte.Text <> "" Then
        sCodPC_F = CStr(CodPCAte.Text)
    Else
        sCodPC_F = ""
    End If

    If sCodPC_I <> "" And sCodPC_F <> "" Then

        If StrParaLong(sCodPC_I) > StrParaLong(sCodPC_F) Then gError 72615

    End If

    'data de Envio inicial não pode ser maior que a final
    If Trim(DataEnvioDe.ClipText) <> "" And Trim(DataEnvioAte.ClipText) <> "" Then
    
         If CDate(DataEnvioDe.Text) > CDate(DataEnvioAte.Text) Then gError 72616
    
    End If
    
    'data inicial não pode ser maior que a final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 72617
    
    End If
    
    'critica Fornecedor Inicial e Final
    If CodFornecedorDe.Text <> "" Then
        sCodFornecedor_I = CStr(CodFornecedorDe.Text)
    Else
        sCodFornecedor_I = ""
    End If

    If CodFornecedorAte.Text <> "" Then
        sCodFornecedor_F = CStr(CodFornecedorAte.Text)
    Else
        sCodFornecedor_F = ""
    End If

    If sCodFornecedor_I <> "" And sCodFornecedor_F <> "" Then

        If StrParaLong(sCodFornecedor_I) > StrParaLong(sCodFornecedor_F) Then gError 72618

    End If

    'critica Fornecedor Inicial e Final
    If NomeFornDe.Text <> "" Then
        sNomeFornecedor_I = CStr(NomeFornDe.Text)
    Else
        sNomeFornecedor_I = ""
    End If

    If NomeFornAte.Text <> "" Then
        sNomeFornecedor_F = CStr(NomeFornAte.Text)
    Else
        sNomeFornecedor_F = ""
    End If

    If sNomeFornecedor_I <> "" And sNomeFornecedor_F <> "" Then

        If sNomeFornecedor_I > sNomeFornecedor_F Then gError 72619

    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 72613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodFilialDe.SetFocus

        Case 72614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeDe.SetFocus

        Case 72615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            CodPCDe.SetFocus

        Case 72616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENVIO_INICIAL_MAIOR", gErr)
            DataEnvioDe.SetFocus

        Case 72617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case 72618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            CodFornecedorDe.SetFocus
    
        Case 72619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            NomeFornDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169296)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodPC_I As String, sCodPC_F As String, sCodFornecedor_I As String, sCodFornecedor_F As String, sNomeFornecedor_I As String, sNomeFornecedor_F As String, sFornecedorTipo As String, sCheckTipo As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao


   If sCodFilial_I <> "" Then sExpressao = "S01"

   If sCodFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S02"

    End If

   If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S03"

    End If

    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S04"

    End If

    If sCodPC_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S05"

    End If

    If sCodPC_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S06"

    End If
    
   If Trim(DataEnvioDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "S07"

    End If

    If Trim(DataEnvioAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S08"

    End If

    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S09"

    End If

    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S10"

    End If

    If sNomeFornecedor_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S11"

    End If

    If sNomeFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S12"

    End If

    If sCodFornecedor_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S13"

    End If

    If sCodFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S14"

    End If

    'Se a opção para apenas um Tipo de Fornecedor estiver selecionada
    If sCheckTipo = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S15"

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169297)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sOrdenacaoPor As String
Dim sTipoFornecedor As String
Dim iTipo As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 72620

    'pega Codigo Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 72621

    CodFilialDe.Text = sParam
    Call CodFilialDe_Validate(bSGECancelDummy)

    'pega  Codigo Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 72622

    CodFilialAte.Text = sParam
    Call CodFilialAte_Validate(bSGECancelDummy)

    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 72623

    NomeDe.Text = sParam
    Call NomeDe_Validate(bSGECancelDummy)

    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 72624

    NomeAte.Text = sParam
    Call NomeAte_Validate(bSGECancelDummy)

    'pega  Codigo PC inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCINIC", sParam)
    If lErro <> SUCESSO Then gError 72625

    CodPCDe.Text = sParam

    'pega  Codigo PC final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCFIM", sParam)
    If lErro <> SUCESSO Then gError 72626

    CodPCAte.Text = sParam

    'pega fornecedor Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 72627

    CodFornecedorDe.Text = sParam

    'pega fornecedor Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 72628

    CodFornecedorAte.Text = sParam

    'pega  Nome do fornecedor Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 72629

    NomeFornDe.Text = sParam
    Call NomeFornDe_Validate(bSGECancelDummy)

    'pega nome do fornecedor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 72630

    NomeFornAte.Text = sParam
    Call NomeFornAte_Validate(bSGECancelDummy)

    'pega DataEnvio inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DENVINIC", sParam)
    If lErro <> SUCESSO Then gError 72631

    Call DateParaMasked(DataEnvioDe, CDate(sParam))

    'pega data de envio final e exibe
    lErro = objRelOpcoes.ObterParametro("DENVFIM", sParam)
    If lErro <> SUCESSO Then gError 72632

    Call DateParaMasked(DataEnvioAte, CDate(sParam))

    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINIC", sParam)
    If lErro <> SUCESSO Then gError 72633

    Call DateParaMasked(DataDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 72634

    Call DateParaMasked(DataAte, CDate(sParam))

    lErro = objRelOpcoes.ObterParametro("NITENS", sParam)
    If lErro <> SUCESSO Then gError 72635

    If sParam = "1" Then
        CheckItens.Value = 1
    Else
        CheckItens.Value = 0
    End If
    
    'pega  Tipo de Fornecedor  e exibe
    lErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lErro <> SUCESSO Then gError 72729
                   
    If sParam = "Todos" Then
    
        Call OptionTodosTipos_Click
        
    Else
    
        'pega  Fornecedor final e exibe
        lErro = objRelOpcoes.ObterParametro("TTFORNECEDOR", sTipoFornecedor)
        If lErro <> SUCESSO Then gError 72730
                        
        OptionUmTipo.Value = True
        ComboTipo.Enabled = True
        ComboTipo.Text = sTipoFornecedor
        Call Combo_Seleciona(ComboTipo, iTipo)
        
    End If
        
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 72636

    Select Case sOrdenacaoPor

            Case "DataEnvio"

                ComboOrdenacao.ListIndex = ORD_POR_DATAENVIO

            Case "Fornecedor"
                
                ComboOrdenacao.ListIndex = ORD_POR_FORNECEDOR

            Case "NomeFornecedor"
                
                ComboOrdenacao.ListIndex = ORD_POR_NOME_FORNECEDOR

            Case "Comprador"
                
                ComboOrdenacao.ListIndex = ORD_POR_COMPRADOR

            Case Else
                gError 72637

    End Select

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 72620 To 72637, 72729, 72730

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169298)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 72638

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 72639

    End If

    Exit Sub

Erro_CodFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 72638

        Case 72639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169299)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 72640

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 72641

    End If

    Exit Sub

Erro_CodFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 72640

        Case 72641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169300)

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
        If lErro <> SUCESSO Then gError 72642

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeDe.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 72643

        NomeDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72642

        Case 72643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169301)

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
        If lErro <> SUCESSO Then gError 72644

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeAte.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 72645

        NomeAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeAte_Validate:

    Cancel = True


    Select Case gErr

        Case 72644

        Case 72645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169302)

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
    Caption = "Relação de Fornecedores de Pedidos de Compra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpFornecedoresPC"

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

        If Me.ActiveControl Is CodPCDe Then
            Call LabelCodPCDe_Click

        ElseIf Me.ActiveControl Is CodPCAte Then
            Call LabelCodPCAte_Click

        ElseIf Me.ActiveControl Is CodFilialDe Then
            Call LabelCodFilialDe_Click

        ElseIf Me.ActiveControl Is CodFilialAte Then
            Call LabelCodFilialAte_Click

        ElseIf Me.ActiveControl Is NomeDe Then
            Call LabelNomeDe_Click

        ElseIf Me.ActiveControl Is NomeAte Then
            Call LabelNomeAte_Click

        ElseIf Me.ActiveControl Is CodFornecedorDe Then
            Call LabelCodFornecedorDe_Click

        ElseIf Me.ActiveControl Is CodFornecedorAte Then
            Call LabelCodFornecedorAte_Click

        ElseIf Me.ActiveControl Is NomeFornDe Then
            Call LabelNomeFornDe_Click

        ElseIf Me.ActiveControl Is NomeFornAte Then
            Call LabelNomeFornAte_Click

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

Private Sub LabelNomeReqDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReqDe, Source, X, Y)
End Sub

Private Sub LabelNomeReqDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReqDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReqAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReqAte, Source, X, Y)
End Sub

Private Sub LabelNomeReqAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReqAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodPCAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodPCAte, Source, X, Y)
End Sub

Private Sub LabelCodPCAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodPCAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodPCDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodPCDe, Source, X, Y)
End Sub

Private Sub LabelCodPCDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodPCDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodFornecedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFornecedorAte, Source, X, Y)
End Sub

Private Sub LabelCodFornecedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFornecedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodFornecedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFornecedorDe, Source, X, Y)
End Sub

Private Sub LabelCodFornecedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFornecedorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornDe, Source, X, Y)
End Sub

Private Sub LabelNomeFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornAte, Source, X, Y)
End Sub

Private Sub LabelNomeFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornAte, Button, Shift, X, Y)
End Sub

