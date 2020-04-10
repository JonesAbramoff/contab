VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl NFeOcx 
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10605
   ScaleHeight     =   6750
   ScaleMode       =   0  'User
   ScaleWidth      =   9716.142
   Begin VB.CheckBox EmContingencia 
      Caption         =   "Em Contingência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2685
      TabIndex        =   68
      Top             =   225
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5235
      Index           =   2
      Left            =   195
      TabIndex        =   0
      Top             =   615
      Visible         =   0   'False
      Width           =   9900
      Begin VB.TextBox NumIntNF 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   8280
         TabIndex        =   61
         Text            =   "NumIntNF"
         Top             =   1335
         Width           =   750
      End
      Begin VB.TextBox Fornecedor 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   7215
         TabIndex        =   60
         Text            =   "Fornecedor"
         Top             =   2895
         Width           =   1755
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   7215
         TabIndex        =   59
         Text            =   "Cliente"
         Top             =   1725
         Width           =   1680
      End
      Begin VB.CommandButton BotaoGerarNFe 
         Caption         =   "Enviar"
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
         Left            =   8205
         TabIndex        =   55
         Top             =   255
         Width           =   1380
      End
      Begin VB.CommandButton ProxNumLote 
         Height          =   285
         Left            =   7770
         Picture         =   "NFeOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Numeração Automática"
         Top             =   255
         Width           =   300
      End
      Begin VB.TextBox Status 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   5250
         TabIndex        =   54
         Text            =   "Status"
         Top             =   3705
         Width           =   4080
      End
      Begin VB.TextBox LoteGrid 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3360
         TabIndex        =   53
         Text            =   "Lote"
         Top             =   3675
         Width           =   1260
      End
      Begin VB.CommandButton BotaoDocOriginal 
         Height          =   690
         Left            =   8040
         Picture         =   "NFeOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4425
         Width           =   1740
      End
      Begin VB.TextBox FilialForn 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   5565
         TabIndex        =   51
         Text            =   "FilialForn"
         Top             =   3150
         Width           =   735
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   225
         Left            =   7185
         TabIndex        =   50
         Top             =   2205
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.TextBox CodFornecedor 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   7080
         TabIndex        =   49
         Text            =   "CodFornecedor"
         Top             =   1095
         Width           =   645
      End
      Begin VB.TextBox TipoNFiscal 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   1560
         TabIndex        =   48
         Text            =   "TipoNFiscal"
         Top             =   2415
         Width           =   960
      End
      Begin VB.TextBox Serie 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1545
         MaxLength       =   50
         TabIndex        =   46
         Top             =   1995
         Width           =   720
      End
      Begin VB.TextBox DataEmissao 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3600
         TabIndex        =   45
         Text            =   "Emissão"
         Top             =   1995
         Width           =   1200
      End
      Begin VB.TextBox FilialCli 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   3720
         TabIndex        =   44
         Text            =   "FilialCli"
         Top             =   2775
         Width           =   675
      End
      Begin VB.TextBox CodCliente 
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   225
         Left            =   7080
         TabIndex        =   43
         Text            =   "CodCliente"
         Top             =   1365
         Width           =   675
      End
      Begin VB.CheckBox Seleciona 
         DragMode        =   1  'Automatic
         Height          =   270
         Left            =   750
         TabIndex        =   42
         Top             =   2475
         Width           =   840
      End
      Begin MSMask.MaskEdBox NumNFiscal 
         Height          =   225
         Left            =   2415
         TabIndex        =   47
         Top             =   1995
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
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
         Left            =   5985
         Picture         =   "NFeOcx.ctx":3000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4485
         Width           =   1650
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
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
         Left            =   4095
         Picture         =   "NFeOcx.ctx":41E2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4485
         Width           =   1650
      End
      Begin VB.ComboBox Ordenados 
         Height          =   315
         ItemData        =   "NFeOcx.ctx":51FC
         Left            =   2055
         List            =   "NFeOcx.ctx":51FE
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   3480
      End
      Begin MSFlexGridLib.MSFlexGrid GridNFiscal 
         Height          =   3390
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   5980
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin MSMask.MaskEdBox Lote 
         Height          =   315
         Left            =   6840
         TabIndex        =   58
         Top             =   255
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
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
      Begin VB.Label LoteLbl 
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
         Left            =   6330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   56
         Top             =   300
         Width           =   450
      End
      Begin VB.Label LabelItensSelecionados 
         Caption         =   "Itens Selecionados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   255
         TabIndex        =   7
         Top             =   4590
         Width           =   1800
      End
      Begin VB.Label Label4 
         Caption         =   "Ordenados por:"
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
         Left            =   510
         TabIndex        =   6
         Top             =   255
         Width           =   1410
      End
      Begin VB.Label ItensSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2070
         TabIndex        =   5
         Top             =   4560
         Width           =   795
      End
   End
   Begin VB.CommandButton BotaoRetCanc 
      Caption         =   "Retorno do Cancelamento de Notas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8820
      TabIndex        =   67
      Top             =   6000
      Width           =   1650
   End
   Begin VB.CommandButton BotaoRetConsulta 
      Caption         =   "Retorno das Consultas de Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   7095
      TabIndex        =   66
      Top             =   6000
      Width           =   1650
   End
   Begin VB.CommandButton BotaoStatusNFe 
      Caption         =   "Status das Notas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5370
      TabIndex        =   65
      Top             =   6000
      Width           =   1650
   End
   Begin VB.CommandButton BotaoRetNFe 
      Caption         =   "Retorno dos Envios de Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3645
      TabIndex        =   64
      Top             =   6000
      Width           =   1650
   End
   Begin VB.CommandButton BotaoLogLote 
      Caption         =   "Log de Envio dos Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1905
      TabIndex        =   63
      Top             =   6000
      Width           =   1650
   End
   Begin VB.CommandButton BotaoLotes 
      Caption         =   "Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   165
      TabIndex        =   62
      Top             =   6000
      Width           =   1650
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5250
      Index           =   1
      Left            =   225
      TabIndex        =   8
      Top             =   600
      Width           =   9885
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Notas Fiscais"
         Height          =   5250
         Left            =   60
         TabIndex        =   9
         Top             =   0
         Width           =   9750
         Begin VB.Frame Frame3 
            Caption         =   "Fornecedores"
            Height          =   795
            Left            =   1995
            TabIndex        =   29
            Top             =   3360
            Width           =   5520
            Begin MSMask.MaskEdBox FornecedorDe 
               Height          =   300
               Left            =   810
               TabIndex        =   30
               Top             =   315
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox FornecedorAte 
               Height          =   300
               Left            =   3465
               TabIndex        =   31
               Top             =   330
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
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
               Left            =   420
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   33
               Top             =   345
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
               Height          =   195
               Left            =   3000
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   32
               Top             =   375
               Width           =   360
            End
         End
         Begin VB.OptionButton Ambas 
            Caption         =   "Ambas"
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
            Left            =   6630
            TabIndex        =   41
            Top             =   375
            Width           =   1455
         End
         Begin VB.OptionButton EnviadasNaoAceitas 
            Caption         =   "Enviadas e Não Aceitas"
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
            Left            =   3660
            TabIndex        =   40
            Top             =   375
            Width           =   2475
         End
         Begin VB.OptionButton NaoEnviadas 
            Caption         =   "Não Enviadas"
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
            Left            =   1665
            TabIndex        =   39
            Top             =   375
            Value           =   -1  'True
            Width           =   1620
         End
         Begin VB.Frame Frame7 
            Caption         =   "Séries"
            Height          =   795
            Left            =   2010
            TabIndex        =   34
            Top             =   765
            Width           =   5520
            Begin VB.ComboBox SerieAte 
               Height          =   315
               Left            =   3465
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   300
               Width           =   1035
            End
            Begin VB.ComboBox SerieDe 
               Height          =   315
               Left            =   810
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   300
               Width           =   1035
            End
            Begin VB.Label Label5 
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
               Left            =   2985
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   36
               Top             =   360
               Width           =   360
            End
            Begin VB.Label Label1 
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
               Left            =   405
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   35
               Top             =   345
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Notas Fiscais"
            Height          =   795
            Left            =   1995
            TabIndex        =   22
            Top             =   1635
            Width           =   5520
            Begin MSMask.MaskEdBox NFiscalDe 
               Height          =   300
               Left            =   810
               TabIndex        =   23
               Top             =   300
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NFiscalAte 
               Height          =   300
               Left            =   3465
               TabIndex        =   24
               Top             =   315
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelNFiscalDe 
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
               Left            =   405
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   26
               Top             =   345
               Width           =   315
            End
            Begin VB.Label LabelNFiscalAte 
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
               Left            =   2985
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   25
               Top             =   360
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Clientes"
            Height          =   795
            Left            =   1995
            TabIndex        =   17
            Top             =   2490
            Width           =   5520
            Begin MSMask.MaskEdBox ClienteDe 
               Height          =   300
               Left            =   810
               TabIndex        =   18
               Top             =   315
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ClienteAte 
               Height          =   300
               Left            =   3465
               TabIndex        =   19
               Top             =   330
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelClienteAte 
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
               Left            =   3000
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   21
               Top             =   375
               Width           =   360
            End
            Begin VB.Label LabelClienteDe 
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
               Left            =   405
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   20
               Top             =   330
               Width           =   315
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data Emissão"
            Height          =   795
            Left            =   1995
            TabIndex        =   10
            Top             =   4230
            Width           =   5520
            Begin MSComCtl2.UpDown UpDownEmissaoAte 
               Height          =   300
               Left            =   4575
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   330
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownEmissaoDe 
               Height          =   300
               Left            =   1920
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   330
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmissaoDe 
               Height          =   300
               Left            =   780
               TabIndex        =   13
               Top             =   330
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataEmissaoAte 
               Height          =   300
               Left            =   3435
               TabIndex        =   14
               Top             =   330
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
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
               Height          =   195
               Left            =   375
               TabIndex        =   16
               Top             =   375
               Width           =   315
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
               Height          =   195
               Left            =   3000
               TabIndex        =   15
               Top             =   390
               Width           =   360
            End
         End
      End
   End
   Begin VB.CommandButton BotaoFechar 
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
      Left            =   9150
      Picture         =   "NFeOcx.ctx":5200
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5700
      Left            =   165
      TabIndex        =   28
      Top             =   225
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   10054
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notas Fiscais"
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
Attribute VB_Name = "NFeOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Event Unload()

Private WithEvents objCT As CTNFe
Attribute objCT.VB_VarHelpID = -1

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


Private Sub BotaoLogLote_Click()
    Call objCT.BotaoLogLote_Click
End Sub

Private Sub BotaoLotes_Click()
    Call objCT.BotaoLotes_Click
End Sub

Private Sub BotaoRetCanc_Click()
    Call objCT.BotaoRetCanc_Click
End Sub

Private Sub BotaoRetConsulta_Click()
    Call objCT.BotaoRetConsulta_Click
End Sub

Private Sub BotaoRetNFe_Click()
    Call objCT.BotaoRetNFe_Click
End Sub

Private Sub BotaoStatusNFe_Click()
    Call objCT.BotaoStatusNFe_Click
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


Private Sub SerieDe_Click()
    Call objCT.SerieDe_Click
End Sub

Private Sub SerieAte_Click()
    Call objCT.SerieAte_Click
End Sub

Private Sub TabStrip1_Click()
    Call objCT.TabStrip1_Click
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

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub LabelNFiscalDe_Click()

     Call objCT.LabelNFiscalDe_Click

End Sub

Private Sub LabelNFiscalAte_Click()

     Call objCT.LabelNFiscalAte_Click

End Sub

Public Sub ClienteDe_Change()

    Call objCT.ClienteDe_Change
    
End Sub

Public Sub ClienteDe_GotFocus()

    Call objCT.ClienteDe_GotFocus
    
End Sub

Public Sub ClienteAte_Change()

    Call objCT.ClienteAte_Change

End Sub

Public Sub ClienteAte_GotFocus()

    Call objCT.ClienteAte_GotFocus
    
End Sub

Public Sub LabelClienteDe_Click()

    Call objCT.LabelClienteDe_Click

End Sub

Public Sub LabelClienteAte_Click()

    Call objCT.LabelClienteAte_Click

End Sub

Public Sub FornecedorDe_Change()

    Call objCT.FornecedorDe_Change
    
End Sub

Public Sub FornecedorDe_GotFocus()

    Call objCT.FornecedorDe_GotFocus
    
End Sub

Public Sub FornecedorAte_Change()

    Call objCT.FornecedorAte_Change

End Sub

Public Sub FornecedorAte_GotFocus()

    Call objCT.FornecedorAte_GotFocus
    
End Sub

Public Sub LabelFornecedorDe_Click()

    Call objCT.LabelFornecedorDe_Click

End Sub

Public Sub LabelFornecedorAte_Click()

    Call objCT.LabelFornecedorAte_Click

End Sub

Public Sub DataEmissaoDe_Change()

    Call objCT.DataEmissaoDe_Change

End Sub

Public Sub DataEmissaoDe_GotFocus()

    Call objCT.DataEmissaoDe_GotFocus

End Sub

Public Sub DataEmissaoDe_Validate(Cancel As Boolean)

    Call objCT.DataEmissaoDe_Validate(Cancel)

End Sub

Public Sub DataEmissaoAte_Change()

    Call objCT.DataEmissaoAte_Change

End Sub

Public Sub DataEmissaoAte_GotFocus()

    Call objCT.DataEmissaoAte_GotFocus

End Sub

Public Sub DataEmissaoAte_Validate(Cancel As Boolean)

    Call objCT.DataEmissaoAte_Validate(Cancel)

End Sub


Public Sub UpDownEmissaoDe_Change()

    Call objCT.UpDownEmissaoDe_Change

End Sub

Public Sub UpDownEmissaoDe_DownClick()

    Call objCT.UpDownEmissaoDe_DownClick

End Sub

Public Sub UpDownEmissaoDe_UpClick()

    Call objCT.UpDownEmissaoDe_UpClick

End Sub

Public Sub UpDownEmissaoAte_Change()

    Call objCT.UpDownEmissaoAte_Change

End Sub

Public Sub UpDownEmissaoAte_DownClick()

    Call objCT.UpDownEmissaoAte_DownClick

End Sub

Public Sub UpDownEmissaoAte_UpClick()

    Call objCT.UpDownEmissaoAte_UpClick

End Sub

Public Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call objCT.TabStrip1_BeforeClick(Cancel)
End Sub

Public Sub Seleciona_Click()
    Call objCT.Seleciona_Click
End Sub

Public Sub Seleciona_GotFocus()
    Call objCT.Seleciona_GotFocus
End Sub

Public Sub Seleciona_KeyPress(KeyAscii As Integer)
    Call objCT.Seleciona_KeyPress(KeyAscii)
End Sub

Public Sub Seleciona_Validate(Cancel As Boolean)
    Call objCT.Seleciona_Validate(Cancel)
End Sub

Public Sub Lote_Change()
    Call objCT.Lote_Change
End Sub

Public Sub Lote_GotFocus()
    Call objCT.Lote_GotFocus
End Sub

Public Sub Ordenados_Change()
    Call objCT.Ordenados_Change
End Sub

Public Sub Ordenados_Click()
    Call objCT.Ordenados_Click
End Sub

Private Sub BotaoGerarNFe_Click()
    Call objCT.BotaoGerarNFe_Click
End Sub

Private Sub ProxNumLote_Click()
    Call objCT.ProxNumLote_Click
End Sub

Public Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub UserControl_Initialize()
    Set objCT = New CTNFe
    Set objCT.objUserControl = Me
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

Public Sub BotaoMarcarTodos_Click()
    Call objCT.BotaoMarcarTodos_Click
End Sub

Public Sub BotaoDesmarcarTodos_Click()
    Call objCT.BotaoDesmarcarTodos_Click
End Sub

Public Sub BotaoDocOriginal_Click()
    Call objCT.BotaoDocOriginal_Click
End Sub

Private Sub BotaoFechar_Click()
    Call objCT.BotaoFechar_Click
End Sub

Private Sub GridNFiscal_Click()
    Call objCT.GridNFiscal_Click
End Sub

Private Sub GridNFiscal_EnterCell()
    Call objCT.GridNFiscal_EnterCell
End Sub

Private Sub GridNFiscal_GotFocus()
    Call objCT.GridNFiscal_GotFocus
End Sub

Private Sub GridNFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.GridNFiscal_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridNFiscal_KeyPress(KeyAscii As Integer)
    Call objCT.GridNFiscal_KeyPress(KeyAscii)
End Sub

Private Sub GridNFiscal_LeaveCell()
    Call objCT.GridNFiscal_LeaveCell
End Sub

Private Sub GridNFiscal_Validate(Cancel As Boolean)
    Call objCT.GridNFiscal_Validate(Cancel)
End Sub

Private Sub GridNFiscal_RowColChange()
    Call objCT.GridNFiscal_RowColChange
End Sub

Private Sub GridNFiscal_Scroll()
    Call objCT.GridNFiscal_Scroll
End Sub

Private Sub NaoEnviadas_Click()
    Call objCT.NaoEnviadas_Click
End Sub

Private Sub EnviadasNaoAceitas_Click()
    Call objCT.EnviadasNaoAceitas_Click
End Sub

Private Sub Ambas_Click()
    Call objCT.Ambas_Click
End Sub

