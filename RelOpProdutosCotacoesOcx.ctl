VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpProdutosCotacoesOcx 
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
   ScaleHeight     =   4620
   ScaleWidth      =   7665
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpProdutosCotacoesOcx.ctx":0000
      Left            =   945
      List            =   "RelOpProdutosCotacoesOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   105
      Width           =   2310
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
      Left            =   3450
      Picture         =   "RelOpProdutosCotacoesOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   135
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5355
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpProdutosCotacoesOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpProdutosCotacoesOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpProdutosCotacoesOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpProdutosCotacoesOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpProdutosCotacoesOcx.ctx":0A9A
      Left            =   1545
      List            =   "RelOpProdutosCotacoesOcx.ctx":0AA4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Frame SSFrame1 
      Caption         =   "Filtros"
      Height          =   3540
      Left            =   180
      TabIndex        =   30
      Top             =   945
      Width           =   6480
      Begin VB.Frame Frame1 
         Caption         =   "Fornecedores"
         Height          =   2775
         Index           =   4
         Left            =   480
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   5535
         Begin VB.Frame Frame11 
            Caption         =   "Código"
            Height          =   705
            Left            =   195
            TabIndex        =   54
            Top             =   465
            Width           =   5100
            Begin MSMask.MaskEdBox FornDe 
               Height          =   300
               Left            =   525
               TabIndex        =   10
               Top             =   285
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox FornAte 
               Height          =   300
               Left            =   2985
               TabIndex        =   11
               Top             =   285
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   8
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoFornAte 
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
               Left            =   2580
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   56
               Top             =   330
               Width           =   360
            End
            Begin VB.Label LabelCodigoFornDe 
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
               TabIndex        =   55
               Top             =   330
               Width           =   315
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Nome Reduzido"
            Height          =   675
            Left            =   195
            TabIndex        =   51
            Top             =   1530
            Width           =   5100
            Begin MSMask.MaskEdBox NomeFornDe 
               Height          =   300
               Left            =   525
               TabIndex        =   12
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeFornAte 
               Height          =   300
               Left            =   3000
               TabIndex        =   13
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
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
               Left            =   2565
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   53
               Top             =   315
               Width           =   360
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
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   52
               Top             =   300
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pedidos de Cotação"
         Height          =   2775
         Index           =   3
         Left            =   480
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   5535
         Begin VB.Frame Frame9 
            Caption         =   "Código"
            Height          =   705
            Left            =   180
            TabIndex        =   48
            Top             =   390
            Width           =   5055
            Begin MSMask.MaskEdBox CodigoPCDe 
               Height          =   300
               Left            =   585
               TabIndex        =   15
               Top             =   270
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoPCAte 
               Height          =   300
               Left            =   3030
               TabIndex        =   16
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
            Begin VB.Label LabelCodigoPCDe 
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
               Left            =   210
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   50
               Top             =   330
               Width           =   315
            End
            Begin VB.Label LabelCodigoPCAte 
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
               Left            =   2625
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   49
               Top             =   300
               Width           =   360
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Data de Validade"
            Height          =   675
            Left            =   180
            TabIndex        =   43
            Top             =   1485
            Width           =   5115
            Begin MSComCtl2.UpDown UpDownDataValidadeDe 
               Height          =   315
               Left            =   1815
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   255
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownDataValidadeAte 
               Height          =   315
               Left            =   4215
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   255
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataValidadeDe 
               Height          =   315
               Left            =   630
               TabIndex        =   17
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
            Begin MSMask.MaskEdBox DataValidadeAte 
               Height          =   315
               Left            =   3030
               TabIndex        =   18
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
            Begin VB.Label Label8 
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
               Left            =   2640
               TabIndex        =   47
               Top             =   315
               Width           =   360
            End
            Begin VB.Label Label7 
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
               Left            =   210
               TabIndex        =   46
               Top             =   315
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtos"
         Height          =   2775
         Index           =   2
         Left            =   480
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   5535
         Begin VB.Frame Frame3 
            Caption         =   "Código"
            Height          =   675
            Left            =   180
            TabIndex        =   40
            Top             =   450
            Width           =   5145
            Begin MSMask.MaskEdBox CodigoProdDe 
               Height          =   300
               Left            =   705
               TabIndex        =   2
               Top             =   255
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodigoProdAte 
               Height          =   300
               Left            =   2385
               TabIndex        =   3
               Top             =   255
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodigoProdAte 
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
               Left            =   1935
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   42
               Top             =   308
               Width           =   360
            End
            Begin VB.Label LabelCodigoProdDe 
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
               TabIndex        =   41
               Top             =   308
               Width           =   315
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Nome Reduzido"
            Height          =   675
            Left            =   180
            TabIndex        =   37
            Top             =   1410
            Width           =   5145
            Begin MSMask.MaskEdBox NomeProdDe 
               Height          =   300
               Left            =   540
               TabIndex        =   4
               Top             =   255
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeProdAte 
               Height          =   300
               Left            =   3060
               TabIndex        =   5
               Top             =   255
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelNomeProdAte 
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
               Left            =   2640
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   39
               Top             =   315
               Width           =   360
            End
            Begin VB.Label LabelNomeProdDe 
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
               TabIndex        =   38
               Top             =   315
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filial Empresa"
         Height          =   2775
         Index           =   1
         Left            =   480
         TabIndex        =   24
         Top             =   600
         Width           =   5535
         Begin VB.Frame FrameNome 
            Caption         =   "Nome"
            Height          =   675
            Left            =   180
            TabIndex        =   34
            Top             =   1440
            Width           =   5160
            Begin MSMask.MaskEdBox NomeFilialDe 
               Height          =   300
               Left            =   540
               TabIndex        =   8
               Top             =   210
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NomeFilialAte 
               Height          =   300
               Left            =   3075
               TabIndex        =   9
               Top             =   210
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
               Left            =   2625
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   36
               Top             =   270
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
               Left            =   165
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   35
               Top             =   270
               Width           =   315
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Código"
            Height          =   660
            Left            =   180
            TabIndex        =   31
            Top             =   405
            Width           =   5145
            Begin MSMask.MaskEdBox CodigoFilialDe 
               Height          =   300
               Left            =   525
               TabIndex        =   6
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
            Begin MSMask.MaskEdBox CodigoFilialAte 
               Height          =   300
               Left            =   3105
               TabIndex        =   7
               Top             =   225
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
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
               Height          =   195
               Left            =   165
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   33
               Top             =   285
               Width           =   315
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
               Left            =   2625
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   32
               Top             =   300
               Width           =   360
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3225
         Left            =   420
         TabIndex        =   57
         Top             =   225
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   5689
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Filiais Empresa"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Produtos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Pedidos de Cotação"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fornecedores"
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
      TabIndex        =   29
      Top             =   135
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
      Left            =   225
      TabIndex        =   28
      Top             =   555
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpProdutosCotacoesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpProdutosCotacoes
Const ORD_POR_CODIGO = 0
Const ORD_POR_DESCRICAO = 1

Private WithEvents objEventoCodigoFornDe As AdmEvento
Attribute objEventoCodigoFornDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFornAte As AdmEvento
Attribute objEventoCodigoFornAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFornDe As AdmEvento
Attribute objEventoNomeFornDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFornAte As AdmEvento
Attribute objEventoNomeFornAte.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFilialDe As AdmEvento
Attribute objEventoCodigoFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFilialAte As AdmEvento
Attribute objEventoCodigoFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoCodProdDe As AdmEvento
Attribute objEventoCodProdDe.VB_VarHelpID = -1
Private WithEvents objEventoCodProdAte As AdmEvento
Attribute objEventoCodProdAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeProdDe As AdmEvento
Attribute objEventoNomeProdDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeProdAte As AdmEvento
Attribute objEventoNomeProdAte.VB_VarHelpID = -1
Private WithEvents objEventoPedCotacaoDe As AdmEvento
Attribute objEventoPedCotacaoDe.VB_VarHelpID = -1
Private WithEvents objEventoPedCotacaoAte As AdmEvento
Attribute objEventoPedCotacaoAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio


Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 74689
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 74690
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 74689
        
        Case 74690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171778)

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
    If lErro <> SUCESSO Then gError 74691
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 74691
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171779)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel
   
End Sub

Private Sub CodigoFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialAte, iAlterado)
    
End Sub

Private Sub CodigoFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialDe, iAlterado)
    
End Sub

Private Sub CodigoPCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPCAte, iAlterado)
    
End Sub

Private Sub CodigoPCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPCDe, iAlterado)
    
End Sub


Private Sub CodigoProdDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdDe_Validate

    If Len(Trim(CodigoProdDe.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74692
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 74693
        
        If lErro = 28030 Then gError 74694
        
    End If
    
    Exit Sub
    
Erro_CodigoProdDe_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 74692, 74693
        
        Case 74694
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171780)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub CodigoProdAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CodigoProdAte_Validate

    If Len(Trim(CodigoProdAte.ClipText)) > 0 Then
        
        lErro = CF("Produto_Formata", CodigoProdAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74695
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 74696
        
        If lErro = 28030 Then gError 74697
        
    End If
    
    Exit Sub
    
Erro_CodigoProdAte_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 74695, 74696
        
        Case 74697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171781)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objEventoCodigoFornDe = New AdmEvento
    Set objEventoCodigoFornAte = New AdmEvento
      
    Set objEventoNomeFornDe = New AdmEvento
    Set objEventoNomeFornAte = New AdmEvento
    
    Set objEventoCodigoFilialDe = New AdmEvento
    Set objEventoCodigoFilialAte = New AdmEvento
    
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
    
    Set objEventoCodProdDe = New AdmEvento
    Set objEventoCodProdAte = New AdmEvento
    
    Set objEventoNomeProdDe = New AdmEvento
    Set objEventoNomeProdAte = New AdmEvento
    
    Set objEventoPedCotacaoDe = New AdmEvento
    Set objEventoPedCotacaoAte = New AdmEvento
    
    'Lê o Tipo de produto do BD
    lErro = CF("Cod_Nomes_Le", "TiposdeProduto", "TipoDeProduto", "Sigla", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 74701
   
    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdDe)
    If lErro <> SUCESSO Then gError 74702

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", CodigoProdAte)
    If lErro <> SUCESSO Then gError 74703

    ComboOrdenacao.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 74700 To 74703
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171782)

    End Select

    Exit Sub

End Sub

Private Sub DataValidadeAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataValidadeAte, iAlterado)
    
End Sub

Private Sub DataValidadeDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataValidadeDe, iAlterado)
    
End Sub

Private Sub FornAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornAte, iAlterado)
    
End Sub

Private Sub FornDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornDe, iAlterado)
    
End Sub

Private Sub FornAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornAte_Validate

    If Len(Trim(FornAte.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornAte.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 74704
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 74705
        
    End If

    Exit Sub

Erro_FornAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74704

        Case 74705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171783)

    End Select

    Exit Sub

End Sub

Private Sub FornDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornDe_Validate

    If Len(Trim(FornDe.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornDe.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 74706
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 74707
        
    End If

    Exit Sub

Erro_FornDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74706

        Case 74707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171784)

    End Select

    Exit Sub

End Sub

Private Sub DataValidadeAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidadeAte_Validate

    'Verifica se a DataValidadeAte está preenchida
    If Len(Trim(DataValidadeAte.Text)) = 0 Then Exit Sub

    'Critica a DataValidadeAte informada
    lErro = Data_Critica(DataValidadeAte.Text)
    If lErro <> SUCESSO Then gError 74805

    Exit Sub
                   
Erro_DataValidadeAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74805
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171785)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoPCDe_Click()

Dim objPedCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_LabelCodigoPCDe_Click

    If Len(Trim(CodigoPCDe.Text)) > 0 Then
    
        objPedCotacao.lCodigo = StrParaLong(CodigoPCDe.Text)
        
    End If
    
    Call Chama_Tela("PedidoCotacaoTodosLista", colSelecao, objPedCotacao, objEventoPedCotacaoDe)
    
    Exit Sub
    
Erro_LabelCodigoPCDe_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171786)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub LabelCodigoPCAte_Click()

Dim objPedCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_LabelCodigoPCAte_Click

    If Len(Trim(CodigoPCAte.Text)) > 0 Then
    
        objPedCotacao.lCodigo = StrParaLong(CodigoPCAte.Text)
        
    End If
    
    Call Chama_Tela("PedidoCotacaoTodosLista", colSelecao, objPedCotacao, objEventoPedCotacaoAte)
    
    Exit Sub
    
Erro_LabelCodigoPCAte_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171787)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoPedCotacaoDe_evSelecao(obj1 As Object)

Dim objPedCotacao As New ClassPedidoCotacao

    Set objPedCotacao = obj1

    CodigoPCDe.Text = CStr(objPedCotacao.lCodigo)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoPedCotacaoAte_evSelecao(obj1 As Object)

Dim objPedCotacao As New ClassPedidoCotacao

    Set objPedCotacao = obj1

    CodigoPCAte.Text = CStr(objPedCotacao.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub UpDownDataValidadeAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidadeAte_DownClick

    'Diminui um dia em DataValidadeAte
    lErro = Data_Up_Down_Click(DataValidadeAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74806

    Exit Sub

Erro_UpDownDataValidadeAte_DownClick:

    Select Case gErr

        Case 74806
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171788)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidadeAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidadeAte_UpClick

    'Diminui um dia em DataValidadeAte
    lErro = Data_Up_Down_Click(DataValidadeAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74807

    Exit Sub

Erro_UpDownDataValidadeAte_UpClick:

    Select Case gErr

        Case 74807
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171789)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidadeDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidadeDe_DownClick

    'Diminui um dia em DataValidadeDe
    lErro = Data_Up_Down_Click(DataValidadeDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74808

    Exit Sub

Erro_UpDownDataValidadeDe_DownClick:

    Select Case gErr

        Case 74808
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171790)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidadeDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataValidadeDe_UpClick

    'Diminui um dia em DataValidadeDe
    lErro = Data_Up_Down_Click(DataValidadeDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74809

    Exit Sub

Erro_UpDownDataValidadeDe_UpClick:

    Select Case gErr

        Case 74809
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171791)

    End Select

    Exit Sub

End Sub

Private Sub DataValidadeDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidadeDe_Validate

    'Verifica se a DataValidadeDe está preenchida
    If Len(Trim(DataValidadeDe.Text)) = 0 Then Exit Sub

    'Critica a DataValidadeDe informada
    lErro = Data_Critica(DataValidadeDe.Text)
    If lErro <> SUCESSO Then gError 74810

    Exit Sub
                   
Erro_DataValidadeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74810
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171792)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodigoFornAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelCodigoFornAte_Click
    
    If Len(Trim(FornAte.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(FornAte.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoCodigoFornAte)

   Exit Sub

Erro_LabelCodigoFornAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171793)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoFornDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelCodigoFornDe_Click
    
    If Len(Trim(FornDe.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(FornDe.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoCodigoFornDe)

   Exit Sub

Erro_LabelCodigoFornDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171794)

    End Select

    Exit Sub

End Sub


Private Sub LabelCodigoProdDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelCodigoProdDe_Click
    
    If Len(Trim(CodigoProdDe.Text)) > 0 Then
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", CodigoProdDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74708
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoCodProdDe)

   Exit Sub

Erro_LabelCodigoProdDe_Click:

    Select Case gErr

        Case 74708
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171795)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoProdAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String

On Error GoTo Erro_LabelCodigoProdAte_Click
    
    If Len(Trim(CodigoProdAte.Text)) > 0 Then
        'Preenche com o Produto da tela
        lErro = CF("Produto_Formata", CodigoProdAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 74709
        
        objProduto.sCodigo = sProdutoFormatado
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoCodProdAte)

   Exit Sub

Erro_LabelCodigoProdAte_Click:

    Select Case gErr

        Case 74709
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171796)

    End Select

    Exit Sub

End Sub
Private Sub LabelNomeProdDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelNomeProdDe_Click
    
    If Len(Trim(NomeProdDe.Text)) > 0 Then
        'Preenche com o Produto da tela
        objProduto.sNomeReduzido = NomeProdDe.Text
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoNomeProdDe)

   Exit Sub

Erro_LabelNomeProdDe_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171797)

    End Select

    Exit Sub
    
End Sub
Private Sub LabelNomeProdAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto

On Error GoTo Erro_LabelNomeProdAte_Click
    
    If Len(Trim(NomeProdAte.Text)) > 0 Then
        'Preenche com o Produto da tela
        objProduto.sNomeReduzido = NomeProdAte.Text
    End If
    
    'Chama Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoNomeProdAte)

   Exit Sub

Erro_LabelNomeProdAte_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171798)

    End Select

    Exit Sub
    
End Sub

Private Sub NomeFornAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NomeFornAte_Validate

    'Verifica se o Nome do Fornecedor foi preenchido
    If Len(Trim(NomeFornAte.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = NomeFornAte.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 74710
        If lErro = 6681 Then gError 74711

    End If
    
    Exit Sub
    
Erro_NomeFornAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 74710
        
        Case 74711
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171799)

    End Select
    
    Exit Sub
    
End Sub
Private Sub NomeFornDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_NomeFornDe_Validate

    'Verifica se o Nome do Fornecedor foi preenchido
    If Len(Trim(NomeFornDe.Text)) > 0 Then
    
        objFornecedor.sNomeReduzido = NomeFornDe.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 74712
        If lErro = 6681 Then gError 74713

    End If
    
    Exit Sub
    
Erro_NomeFornDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 74712
        
        Case 74713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171800)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCodigoFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    FornAte.Text = CStr(objFornecedor.lCodigo)
    Call FornAte_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    FornDe.Text = CStr(objFornecedor.lCodigo)
    Call FornDe_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeFilialDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objFilialEmpresa.sNome = NomeFilialDe.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171801)

    End Select

    Exit Sub

End Sub
Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.sNome = NomeFilialAte.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171802)

    End Select

    Exit Sub

End Sub
Private Sub NomeFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeFilialDe_Validate

    bAchou = False
    
    If Len(Trim(NomeFilialDe.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 74714

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeFilialDe.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74715
        
        NomeFilialDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74714

        Case 74715
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171803)

    End Select

Exit Sub

End Sub

Private Sub NomeFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeFilialAte_Validate

    bAchou = False
    If Len(Trim(NomeFilialAte.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 74716

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeFilialAte.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74717

        NomeFilialAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74716

        Case 74717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171804)

    End Select

Exit Sub

End Sub

Private Sub CodigoFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialDe_Validate

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then

        'Lê o código informado
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74718
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74719

    End If

    Exit Sub

Erro_CodigoFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 74718

        Case 74719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.lCodEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171805)

    End Select

    Exit Sub

End Sub
Private Sub CodigoFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialAte_Validate

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74720
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74721

    End If

    Exit Sub

Erro_CodigoFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74720

        Case 74721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.lCodEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171806)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoDe_Click

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.lCodEmpresa = StrParaLong(CodigoFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoFilialDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171807)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodigoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoAte_Click

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.lCodEmpresa = StrParaLong(CodigoFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoFilialAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171808)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeFornDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNomeFornDe_Click
    
    If Len(Trim(NomeFornDe.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.sNomeReduzido = NomeFornDe.Text
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoNomeFornDe)

   Exit Sub

Erro_LabelNomeFornDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171809)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeFornAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNomeFornAte_Click
    
    If Len(Trim(NomeFornAte.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.sNomeReduzido = NomeFornAte.Text
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoNomeFornAte)

   Exit Sub

Erro_LabelNomeFornAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171810)

    End Select

    Exit Sub

End Sub
Private Sub objEventoCodProdAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCodProdAte_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 74722
    
    CodigoProdAte.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoCodProdAte_evSelecao:

    Select Case gErr
    
        Case 74722
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171811)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub objEventoCodProdDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCodProdDe_evSelecao

    Set objProduto = obj1

    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 74723
    
    CodigoProdDe.Text = sProdutoMascarado

    Me.Show

    Exit Sub

Erro_objEventoCodProdDe_evSelecao:

    Select Case gErr
    
        Case 74723
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171812)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoNomeFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialAte.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialDe.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    NomeFornDe.Text = objFornecedor.sNomeReduzido
    Call NomeFornDe_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoNomeFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    NomeFornAte.Text = objFornecedor.sNomeReduzido
    Call NomeFornAte_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeProdDe_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto

    Set objProduto = obj1
    
    NomeProdDe.Text = objProduto.sNomeReduzido

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoNomeProdAte_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto

    Set objProduto = obj1
    
    NomeProdAte.Text = objProduto.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 74724

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74725

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 74726
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 74727
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 74724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 74725, 74726, 74727
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171813)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 74728

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 74729

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 74728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 74729

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171814)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74730

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEmissao", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CondPagto", 1)
                
                
            Case ORD_POR_DESCRICAO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoDesc", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ProdutoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEmissao", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CondPagto", 1)
                
            Case Else
                gError 74963

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 74730, 74963

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171815)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sFornecedor_I As String
Dim sFornecedor_F As String
Dim sNomeForn_I As String
Dim sNomeForn_F As String
Dim sFilial_I As String
Dim sFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sOrdenacaoPor As String
Dim sCheckTipo As String
Dim sFornecedorTipo As String
Dim sNomeProd_I As String
Dim sNomeProd_F As String
Dim sCodProd_I As String
Dim sCodProd_F As String
Dim sCodPC_I As String
Dim sCodPC_F As String
Dim sCheckTipoProd As String
Dim sProdutoTipo As String
Dim sNatureza_I As String
Dim sNatureza_F As String
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp
 
    lErro = Formata_E_Critica_Parametros(sFornecedor_I, sFornecedor_F, sNomeForn_I, sNomeForn_F, sFilial_I, sFilial_F, sNomeFilial_I, sNomeFilial_F, sCodProd_I, sCodProd_F, sNomeProd_I, sNomeProd_F, sNatureza_I, sNatureza_F, sCheckTipo, sFornecedorTipo, sCheckTipoProd, sProdutoTipo, sCodPC_I, sCodPC_F)
    If lErro <> SUCESSO Then gError 74731

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 74732
         
    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then gError 74733
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNINIC", NomeFornDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74734
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 74735
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeFilialDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74736
    
    lErro = objRelOpcoes.IncluirParametro("TCODPRODINIC", sCodProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 74737
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEPRODINIC", NomeProdDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74738
    
    lErro = objRelOpcoes.IncluirParametro("TNATUREZAPRODINIC", sNatureza_I)
    If lErro <> AD_BOOL_TRUE Then gError 74739
    
    lErro = objRelOpcoes.IncluirParametro("NCODPEDCOTINIC", sCodPC_I)
    If lErro <> AD_BOOL_TRUE Then gError 74740
    
    'Preenche a data validade inicial
    If Trim(DataValidadeDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVALIDADEINIC", DataValidadeDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVALIDADEINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74742
    
    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then gError 74743
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFORNFIM", NomeFornAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74744
        
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 74745
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeFilialAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74746
        
    lErro = objRelOpcoes.IncluirParametro("TCODPRODFIM", sCodProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 74747
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEPRODFIM", NomeProdAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74748
    
    lErro = objRelOpcoes.IncluirParametro("TNATUREZAPRODFIM", sNatureza_F)
    If lErro <> AD_BOOL_TRUE Then gError 74749
        
    lErro = objRelOpcoes.IncluirParametro("NCODPEDCOTFIM", sCodPC_F)
    If lErro <> AD_BOOL_TRUE Then gError 74750
    
    'Preenche a data validade inicial
    If Trim(DataValidadeAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVALIDADEFIM", DataValidadeAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVALIDADEFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74752
        
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "Codigo"
                
            Case ORD_POR_DESCRICAO
                
                sOrdenacaoPor = "Descricao"
                
            Case Else
                gError 74753
                  
    End Select
        
    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 74754

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 74851
    
    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 74756

    lErro = objRelOpcoes.IncluirParametro("TOPTIPOP", sCheckTipoProd)
    If lErro <> AD_BOOL_TRUE Then gError 74758

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFornecedor_I, sFornecedor_F, sNomeForn_I, sNomeForn_F, sFilial_I, sFilial_F, sNomeFilial_I, sNomeFilial_F, sNomeProd_I, sNomeProd_F, sCodProd_I, sCodProd_F, sNatureza_I, sNatureza_F, sFornecedorTipo, sCheckTipo, sOrdenacaoPor, sCheckTipoProd, sProdutoTipo, sCodPC_I, sCodPC_F, sOrd)
    If lErro <> SUCESSO Then gError 74759

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 74731 To 74759, 74851
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171816)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sFornecedor_I As String, sFornecedor_F As String, sNomeForn_I As String, sNomeForn_F As String, sFilial_I As String, sFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodProd_I As String, sCodProd_F As String, sNomeProd_I As String, sNomeProd_F As String, sNatureza_I As String, sNatureza_F As String, sCheckTipo As String, sFornecedorTipo As String, sCheckProd As String, sProdutoTipo As String, sCodPC_I As String, sCodPC_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Fornecedor Inicial e Final
    If FornDe.Text <> "" Then
        sFornecedor_I = CStr(FornDe.Text)
    Else
        sFornecedor_I = ""
    End If
    
    If FornAte.Text <> "" Then
        sFornecedor_F = CStr(FornAte.Text)
    Else
        sFornecedor_F = ""
    End If
            
    If sFornecedor_I <> "" And sFornecedor_F <> "" Then
        
        If CLng(sFornecedor_I) > CLng(sFornecedor_F) Then gError 74760
        
    End If
                
    'critica NomeFornecedor Inicial e Final
    If NomeFornDe.Text <> "" Then
        sNomeForn_I = NomeFornDe.Text
    Else
        sNomeForn_I = ""
    End If
    
    If NomeFornAte.Text <> "" Then
        sNomeForn_F = NomeFornAte.Text
    Else
        sNomeForn_F = ""
    End If
            
    If sNomeForn_I <> "" And sNomeForn_F <> "" Then
        
        If sNomeForn_I > sNomeForn_F Then gError 74761
        
    End If
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", CodigoProdDe.Text, sCodProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 74989

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sCodProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", CodigoProdAte.Text, sCodProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 74990

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sCodProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sCodProd_I > sCodProd_F Then gError 74762

    End If
    
                
    'critica Nome Produto Inicial e Final
    If NomeProdDe.Text <> "" Then
        sNomeProd_I = NomeProdDe.Text
    Else
        sNomeProd_I = ""
    End If
    
    If NomeProdAte.Text <> "" Then
        sNomeProd_F = NomeProdAte.Text
    Else
        sNomeProd_F = ""
    End If
            
    If sNomeProd_I <> "" And sNomeProd_F <> "" Then
        
        If sNomeProd_I > sNomeProd_F Then gError 74763
        
    End If
    
    'critica Filial Inicial e Final
    If CodigoFilialDe.Text <> "" Then
        sFilial_I = CStr(CodigoFilialDe.Text)
    Else
        sFilial_I = ""
    End If
    
    If CodigoFilialAte.Text <> "" Then
        sFilial_F = CStr(CodigoFilialAte.Text)
    Else
        sFilial_F = ""
    End If
            
    If sFilial_I <> "" And sFilial_F <> "" Then
        
        If CLng(sFilial_I) > CLng(sFilial_F) Then gError 74764
        
    End If
    
    'critica NomeFilial Inicial e Final
    If NomeFilialDe.Text <> "" Then
        sNomeFilial_I = NomeFilialDe.Text
    Else
        sNomeFilial_I = ""
    End If
    
    If NomeFilialAte.Text <> "" Then
        sNomeFilial_F = NomeFilialAte.Text
    Else
        sNomeFilial_F = ""
    End If
            
    If sNomeFilial_I <> "" And sNomeFilial_F <> "" Then
        
        If sNomeFilial_I > sNomeFilial_F Then gError 74765
        
    End If
      
    'critica PedidoCotacao Inicial e Final
    If CodigoPCDe.Text <> "" Then
        sCodPC_I = CStr(CodigoPCDe.Text)
    Else
        sCodPC_I = ""
    End If
    
    If CodigoPCAte.Text <> "" Then
        sCodPC_F = CStr(CodigoPCAte.Text)
    Else
        sCodPC_F = ""
    End If
            
    If sCodPC_I <> "" And sCodPC_F <> "" Then
        
        If CLng(sCodPC_I) > CLng(sCodPC_F) Then gError 74769
        
    End If
        
    'data de Validade inicial não pode ser maior que a data  final
    If Trim(DataValidadeDe.ClipText) <> "" And Trim(DataValidadeAte.ClipText) <> "" Then
    
         If CDate(DataValidadeDe.Text) > CDate(DataValidadeAte.Text) Then gError 74771
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 74760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornDe.SetFocus
                
        Case 74761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            NomeFornDe.SetFocus
            
        Case 74762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            CodigoProdDe.SetFocus
            
        Case 74763
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            NomeProdDe.SetFocus
            
        Case 74764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodigoFilialDe.SetFocus
            
        Case 74765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeFilialDe.SetFocus
        
        Case 74767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 74768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 74769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDCOTACAO_INICIAL_MAIOR", gErr)
            CodigoPCDe.SetFocus
            
        Case 74771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVALIDADE_INICIAL_MAIOR", gErr)
            
        Case 74989
            CodigoProdDe.SetFocus
            
        Case 74990
            CodigoProdAte.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171817)

    End Select

    Exit Function

End Function

                                                                        
Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFornecedor_I As String, sFornecedor_F As String, sNomeForn_I As String, sNomeForn_F As String, sFilial_I As String, sFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sNomeProd_I As String, sNomeProd_F As String, sCodProd_I As String, sCodProd_F As String, sNatureza_I As String, sNatureza_F As String, sFornecedorTipo As String, sCheckTipo As String, sOrdenacaoPor As String, sCheckProd As String, sProdutoTipo As String, sCodPC_I As String, sCodPC_F As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sFornecedor_I <> "" Then sExpressao = "CodForn >= " & Forprint_ConvLong(CLng(sFornecedor_I))

    If sFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodForn <= " & Forprint_ConvLong(CLng(sFornecedor_F))

    End If
           
    If sNomeForn_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NomeFornInic"

    End If

    If sNomeForn_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NomeFornFim"

    End If
    
    If sNomeProd_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NomeProdInic"

    End If
    
    If sNomeProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NomeProdFim"

    End If
    
    If sCodProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodProd >= " & Forprint_ConvTexto(sCodProd_I)

    End If
    
    If sCodProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodProd <= " & Forprint_ConvTexto(sCodProd_F)

    End If
    
    If sCodPC_I <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodPedCot >= " & Forprint_ConvLong(CLng(sCodPC_I))

    End If
   
    If sCodPC_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodPedCot <= " & Forprint_ConvLong(CLng(sCodPC_F))

    End If
       
    If sFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCod >= " & Forprint_ConvInt(StrParaInt(sFilial_I))

    End If
    
    If sFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCod <= " & Forprint_ConvInt(StrParaInt(sFilial_F))

    End If
           
    If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeInic"

    End If
    
    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeFim"

    End If
        
    If Trim(DataValidadeDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ValInic"

    End If
    
    If Trim(DataValidadeAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ValFim"

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171818)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sTipoFornecedor As String, iTipo As Integer
Dim sOrdenacaoPor As String
Dim sTipoProduto As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 74772
   
    'pega Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 74773
    
    FornDe.Text = sParam
    Call FornDe_Validate(bSGECancelDummy)
    
    'pega  Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 74774
    
    FornAte.Text = sParam
    Call FornAte_Validate(bSGECancelDummy)
                                
    'pega Nome do Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 74775
    
    NomeFornDe.Text = sParam
    Call NomeFornDe_Validate(bSGECancelDummy)
    
    'pega  Nome do Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 74776
    
    NomeFornAte.Text = sParam
    Call NomeFornAte_Validate(bSGECancelDummy)
                            
    'pega Nome do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 74777
    
    NomeProdDe.Text = sParam
    
    'pega  Nome do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 74778
    
    NomeProdAte.Text = sParam
    
    'pega codigo do produto inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCODPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 74779
    
    CodigoProdDe.PromptInclude = False
    CodigoProdDe.Text = sParam
    CodigoProdDe.PromptInclude = True
    
    Call CodigoProdDe_Validate(bSGECancelDummy)
    
    'pega  codigo do produto final e exibe
    lErro = objRelOpcoes.ObterParametro("TCODPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 74780
    
    CodigoProdAte.PromptInclude = False
    CodigoProdAte.Text = sParam
    CodigoProdAte.PromptInclude = True
    
    Call CodigoProdAte_Validate(bSGECancelDummy)
    
    'pega Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74783
    
    CodigoFilialDe.Text = sParam
    Call FornDe_Validate(bSGECancelDummy)
    
    'pega  Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74784
    
    CodigoFilialAte.Text = sParam
    Call CodigoFilialAte_Validate(bSGECancelDummy)
                                
    'pega Nome da Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74785
    
    NomeFilialDe.Text = sParam
    Call NomeFilialDe_Validate(bSGECancelDummy)
    
    'pega  Nome da Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74786
    
    NomeFilialAte.Text = sParam
    Call NomeFilialAte_Validate(bSGECancelDummy)
                    
    'pega CodigoPC inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPEDCOTINIC", sParam)
    If lErro <> SUCESSO Then gError 74791
    
    CodigoPCDe.Text = sParam
    
    'pega CodigoPC final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPEDCOTFIM", sParam)
    If lErro <> SUCESSO Then gError 74792
    
    CodigoPCAte.Text = sParam
    
    'pega data validade inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DVALIDADEINIC", sParam)
    If lErro <> SUCESSO Then gError 74795

    Call DateParaMasked(DataValidadeDe, CDate(sParam))
       
    'pega data validade final e exibe
    lErro = objRelOpcoes.ObterParametro("DVALIDADEFIM", sParam)
    If lErro <> SUCESSO Then gError 74796
    
    Call DateParaMasked(DataValidadeAte, CDate(sParam))
    
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 74797
    
    Select Case sOrdenacaoPor
        
            Case "Codigo"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "Descricao"
            
                ComboOrdenacao.ListIndex = ORD_POR_DESCRICAO
                                            
            Case Else
                gError 74798
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 74772 To 74798
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171819)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCodigoFornDe = Nothing
    Set objEventoCodigoFornAte = Nothing
    
    Set objEventoNomeFornDe = Nothing
    Set objEventoNomeFornAte = Nothing
    
    Set objEventoCodigoFilialDe = Nothing
    Set objEventoCodigoFilialAte = Nothing
    
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
    
    Set objEventoNomeProdDe = Nothing
    Set objEventoNomeProdAte = Nothing
    
    Set objEventoCodProdDe = Nothing
    Set objEventoCodProdAte = Nothing
    
    Set objEventoPedCotacaoDe = Nothing
    Set objEventoPedCotacaoAte = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CADFORN
    Set Form_Load_Ocx = Me
    Caption = "Produtos x Cotações"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpProdutosCotacoes"
    
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


Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
            
            'Esconde o frame atual, mostra o novo
            Frame1(TabStrip1.SelectedItem.Index).Visible = True
            Frame1(iFrameAtual).Visible = False
            'Armazena novo valor de iFrameAtual
            iFrameAtual = TabStrip1.SelectedItem.Index

        End If
        
    
    Exit Sub

Erro_TabStrip1_Click:
    
    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171820)

    End Select

    Exit Sub

End Sub

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
        
        If Me.ActiveControl Is FornDe Then
            Call LabelCodigoFornDe_Click
        ElseIf Me.ActiveControl Is FornAte Then
            Call LabelCodigoFornAte_Click
        ElseIf Me.ActiveControl Is NomeFornDe Then
            Call LabelNomeFornDe_Click
        ElseIf Me.ActiveControl Is NomeFornAte Then
            Call LabelNomeFornAte_Click
        ElseIf Me.ActiveControl Is CodigoFilialDe Then
            Call LabelCodigoDe_Click
        ElseIf Me.ActiveControl Is CodigoFilialAte Then
            Call LabelCodigoAte_Click
        ElseIf Me.ActiveControl Is NomeFilialDe Then
            Call LabelNomeDe_Click
        ElseIf Me.ActiveControl Is NomeFilialAte Then
            Call LabelNomeAte_Click
        ElseIf Me.ActiveControl Is CodigoProdDe Then
            Call LabelCodigoProdDe_Click
        ElseIf Me.ActiveControl Is CodigoProdAte Then
            Call LabelCodigoProdAte_Click
        ElseIf Me.ActiveControl Is NomeProdDe Then
            Call LabelNomeProdDe_Click
        ElseIf Me.ActiveControl Is NomeProdAte Then
            Call LabelNomeProdAte_Click
        ElseIf Me.ActiveControl Is CodigoPCDe Then
            Call LabelCodigoPCDe_Click
        ElseIf Me.ActiveControl Is CodigoPCAte Then
            Call LabelCodigoPCAte_Click
        End If
    
    End If

End Sub



Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub






Private Sub LabelCodigoFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoFornAte, Source, X, Y)
End Sub

Private Sub LabelCodigoFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoFornAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoFornDe, Source, X, Y)
End Sub

Private Sub LabelCodigoFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoFornDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornAte, Source, X, Y)
End Sub

Private Sub LabelNomeFornAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeFornDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeFornDe, Source, X, Y)
End Sub

Private Sub LabelNomeFornDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeFornDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoPCDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoPCDe, Source, X, Y)
End Sub

Private Sub LabelCodigoPCDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoPCDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoPCAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoPCAte, Source, X, Y)
End Sub

Private Sub LabelCodigoPCAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoPCAte, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdAte, Source, X, Y)
End Sub

Private Sub LabelCodigoProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoProdDe, Source, X, Y)
End Sub

Private Sub LabelCodigoProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoProdDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeProdAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeProdAte, Source, X, Y)
End Sub

Private Sub LabelNomeProdAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeProdAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeProdDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeProdDe, Source, X, Y)
End Sub

Private Sub LabelNomeProdDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeProdDe, Button, Shift, X, Y)
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

Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
End Sub

