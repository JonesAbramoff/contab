VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl EtapaPRJ 
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   9510
   Begin VB.Frame FrameVitorias 
      BorderStyle     =   0  'None
      Caption         =   "FrameVitorias"
      Height          =   705
      Left            =   135
      TabIndex        =   307
      Top             =   -75
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame Frame12 
         Caption         =   "Última vistoria cadastrada"
         Height          =   615
         Left            =   2940
         TabIndex        =   310
         Top             =   60
         Width           =   4110
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Validade:"
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
            Left            =   2055
            TabIndex        =   314
            Top             =   255
            Width           =   810
         End
         Begin VB.Label ValidadeVistoria 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2925
            TabIndex        =   313
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Left            =   150
            TabIndex        =   312
            Top             =   255
            Width           =   480
         End
         Begin VB.Label DataVistoria 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   660
            TabIndex        =   311
            Top             =   210
            Width           =   1095
         End
      End
      Begin VB.CommandButton BotaoConVistorias 
         Caption         =   "Consultar Vistorias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1485
         TabIndex        =   309
         Top             =   135
         Width           =   1395
      End
      Begin VB.CommandButton BotaoCadVistorias 
         Caption         =   "Cadastrar Vistorias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   45
         TabIndex        =   308
         Top             =   135
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4875
      Index           =   1
      Left            =   135
      TabIndex        =   65
      Top             =   1200
      Width           =   9285
      Begin VB.Frame Frame8 
         Caption         =   "Dados do Cliente"
         Height          =   630
         Left            =   90
         TabIndex        =   147
         Top             =   2190
         Width           =   8670
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5895
            TabIndex        =   6
            Top             =   180
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1710
            TabIndex        =   5
            Top             =   210
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            Index           =   13
            Left            =   5340
            TabIndex        =   149
            Top             =   255
            Width           =   465
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
            Left            =   975
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   148
            Top             =   255
            Width           =   660
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Identificação"
         Height          =   1995
         Left            =   90
         TabIndex        =   71
         Top             =   165
         Width           =   8670
         Begin VB.TextBox Descricao 
            Height          =   300
            Left            =   1680
            MaxLength       =   255
            TabIndex        =   4
            Top             =   1515
            Width           =   6870
         End
         Begin MSMask.MaskEdBox NomeReduzido 
            Height          =   315
            Left            =   5865
            TabIndex        =   3
            Top             =   1125
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Referencia 
            Height          =   300
            Left            =   1680
            TabIndex        =   2
            Top             =   750
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   315
            Left            =   1680
            TabIndex        =   0
            Top             =   315
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReduzidoPRJ 
            Height          =   315
            Left            =   5865
            TabIndex        =   1
            Top             =   270
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeRedPRJ 
            Caption         =   "Nome Projeto:"
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
            Height          =   315
            Left            =   4590
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   146
            Top             =   315
            Width           =   1230
         End
         Begin VB.Label LabelProjeto 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   945
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   145
            Top             =   360
            Width           =   675
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   705
            TabIndex        =   137
            Top             =   1560
            Width           =   930
         End
         Begin VB.Label Codigo 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   136
            Top             =   1125
            Width           =   2040
         End
         Begin VB.Label LabelCodigo 
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
            Left            =   975
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   135
            Top             =   1170
            Width           =   660
         End
         Begin VB.Label LabelReferencia 
            AutoSize        =   -1  'True
            Caption         =   "Referência:"
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
            Left            =   630
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   73
            Top             =   795
            Width           =   1005
         End
         Begin VB.Label LabelNomeReduzido 
            Caption         =   "Nome Reduzido:"
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
            Height          =   315
            Left            =   4380
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   72
            Top             =   1155
            Width           =   1410
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Outros"
         Height          =   1965
         Left            =   90
         TabIndex        =   66
         Top             =   2835
         Width           =   8670
         Begin VB.CheckBox RespUsu 
            Caption         =   "Usuário do Sistema"
            Enabled         =   0   'False
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
            Left            =   5130
            TabIndex        =   8
            Top             =   315
            Width           =   2400
         End
         Begin VB.ComboBox Responsavel 
            Height          =   315
            Left            =   1710
            TabIndex        =   7
            Top             =   270
            Width           =   3210
         End
         Begin VB.TextBox Observacao 
            Height          =   330
            Left            =   1725
            MaxLength       =   255
            TabIndex        =   11
            Top             =   1500
            Width           =   6825
         End
         Begin VB.TextBox Objetivo 
            Height          =   330
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   9
            Top             =   675
            Width           =   6810
         End
         Begin VB.TextBox Justificativa 
            Height          =   330
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   10
            Top             =   1080
            Width           =   6825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Responsável:"
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
            Index           =   11
            Left            =   465
            TabIndex        =   70
            Top             =   315
            Width           =   1170
         End
         Begin VB.Label Label1 
            Caption         =   "Observação:"
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
            Index           =   5
            Left            =   510
            TabIndex        =   69
            Top             =   1590
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Objetivo:"
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
            Height          =   330
            Index           =   4
            Left            =   840
            TabIndex        =   68
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Justificativa:"
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
            Height          =   330
            Index           =   6
            Left            =   510
            TabIndex        =   67
            Top             =   1140
            Width           =   1110
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Dados Complementares"
      Height          =   4875
      Index           =   9
      Left            =   135
      TabIndex        =   150
      Top             =   1200
      Visible         =   0   'False
      Width           =   9285
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   3990
         MaxLength       =   255
         TabIndex        =   46
         Top             =   1725
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   3990
         MaxLength       =   255
         TabIndex        =   45
         Top             =   1365
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3990
         MaxLength       =   255
         TabIndex        =   43
         Top             =   660
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   3990
         MaxLength       =   255
         TabIndex        =   44
         Top             =   1005
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   3990
         MaxLength       =   255
         TabIndex        =   42
         Top             =   300
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.ComboBox Controles 
         Height          =   315
         ItemData        =   "EtapaPRJ.ctx":0000
         Left            =   7230
         List            =   "EtapaPRJ.ctx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   4410
         Width           =   1770
      End
      Begin VB.CommandButton BotaoDadosCustNovo 
         Height          =   405
         Left            =   6270
         Picture         =   "EtapaPRJ.ctx":0030
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   4365
         Width           =   435
      End
      Begin VB.CommandButton BotaoDadosCustDel 
         Height          =   405
         Left            =   6765
         Picture         =   "EtapaPRJ.ctx":0542
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4350
         Width           =   435
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   1
         Left            =   2475
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   1
         Left            =   1305
         TabIndex        =   32
         Top             =   345
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   2
         Left            =   2475
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   2
         Left            =   1305
         TabIndex        =   34
         Top             =   705
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   3
         Left            =   2475
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   3
         Left            =   1305
         TabIndex        =   36
         Top             =   1065
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   4
         Left            =   2475
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1455
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   4
         Left            =   1305
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   5
         Left            =   2475
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1815
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   5
         Left            =   1305
         TabIndex        =   40
         Top             =   1800
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   1
         Left            =   1305
         TabIndex        =   47
         Top             =   2265
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   2
         Left            =   1305
         TabIndex        =   48
         Top             =   2625
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   3
         Left            =   1305
         TabIndex        =   49
         Top             =   2985
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   4
         Left            =   1305
         TabIndex        =   50
         Top             =   3345
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Index           =   5
         Left            =   1305
         TabIndex        =   51
         Top             =   3720
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   1
         Left            =   3990
         TabIndex        =   52
         Top             =   2265
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   2
         Left            =   3990
         TabIndex        =   53
         Top             =   2625
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   3
         Left            =   3990
         TabIndex        =   54
         Top             =   2985
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   4
         Left            =   3990
         TabIndex        =   55
         Top             =   3345
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Index           =   5
         Left            =   3990
         TabIndex        =   56
         Top             =   3720
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data1:"
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
         Index           =   1001
         Left            =   165
         TabIndex        =   201
         Top             =   420
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data2:"
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
         Index           =   1002
         Left            =   165
         TabIndex        =   200
         Top             =   765
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data3:"
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
         Index           =   1003
         Left            =   165
         TabIndex        =   199
         Top             =   1095
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data4:"
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
         Index           =   1004
         Left            =   165
         TabIndex        =   198
         Top             =   1485
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data5:"
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
         Index           =   1005
         Left            =   165
         TabIndex        =   197
         Top             =   1845
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto1:"
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
         Index           =   4001
         Left            =   2940
         TabIndex        =   196
         Top             =   405
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto2:"
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
         Index           =   4002
         Left            =   3255
         TabIndex        =   195
         Top             =   765
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto3:"
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
         Index           =   4003
         Left            =   3255
         TabIndex        =   194
         Top             =   1095
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto4:"
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
         Index           =   4004
         Left            =   3255
         TabIndex        =   193
         Top             =   1485
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto5:"
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
         Index           =   4005
         Left            =   3255
         TabIndex        =   192
         Top             =   1845
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número1:"
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
         Height          =   285
         Index           =   3001
         Left            =   120
         TabIndex        =   191
         Top             =   2310
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número2:"
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
         Height          =   285
         Index           =   3002
         Left            =   120
         TabIndex        =   190
         Top             =   2670
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número3:"
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
         Height          =   285
         Index           =   3003
         Left            =   120
         TabIndex        =   189
         Top             =   3030
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número4:"
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
         Height          =   285
         Index           =   3004
         Left            =   120
         TabIndex        =   188
         Top             =   3390
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número5:"
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
         Height          =   285
         Index           =   3005
         Left            =   120
         TabIndex        =   187
         Top             =   3765
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor1:"
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
         Index           =   2001
         Left            =   2640
         TabIndex        =   186
         Top             =   2325
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor2:"
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
         Index           =   2002
         Left            =   2640
         TabIndex        =   185
         Top             =   2700
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor3:"
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
         Index           =   2003
         Left            =   2640
         TabIndex        =   184
         Top             =   3090
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor4:"
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
         Index           =   2004
         Left            =   2640
         TabIndex        =   183
         Top             =   3450
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor5:"
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
         Index           =   2005
         Left            =   2640
         TabIndex        =   182
         Top             =   3810
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Mao de Obra"
      Height          =   4875
      Index           =   6
      Left            =   135
      TabIndex        =   106
      Top             =   1200
      Visible         =   0   'False
      Width           =   9285
      Begin VB.Frame FrameMO 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   3
         Left            =   240
         TabIndex        =   221
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MOOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   294
            Top             =   1905
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MOData 
            Height          =   315
            Index           =   3
            Left            =   5250
            TabIndex        =   293
            Top             =   2235
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoMOTrazer 
            Caption         =   "Trazer Calculado"
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
            Index           =   3
            Left            =   1875
            TabIndex        =   223
            Top             =   3960
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MOCustoT 
            Height          =   315
            Index           =   3
            Left            =   2955
            TabIndex        =   259
            Top             =   1530
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MODescricao 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   3885
            MaxLength       =   50
            TabIndex        =   225
            Top             =   735
            Width           =   2235
         End
         Begin VB.TextBox MOCodigo 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   435
            MaxLength       =   20
            TabIndex        =   224
            Top             =   795
            Width           =   1245
         End
         Begin VB.CommandButton BotaoMO 
            Caption         =   "Mão de Obra"
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
            Index           =   3
            Left            =   75
            TabIndex        =   222
            Top             =   3960
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MOQuantidade 
            Height          =   315
            Index           =   3
            Left            =   5700
            TabIndex        =   226
            Top             =   315
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOHoras 
            Height          =   315
            Index           =   3
            Left            =   6285
            TabIndex        =   227
            Top             =   765
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOCusto 
            Height          =   315
            Index           =   3
            Left            =   3840
            TabIndex        =   228
            Top             =   105
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMO 
            Height          =   3075
            Index           =   3
            Left            =   75
            TabIndex        =   229
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
         Begin VB.Label MOCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   7155
            TabIndex        =   274
            Top             =   3960
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   29
            Left            =   6075
            TabIndex        =   273
            Top             =   4020
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMO 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   2
         Left            =   240
         TabIndex        =   107
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MOOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   4275
            MaxLength       =   50
            TabIndex        =   304
            Top             =   2565
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MOCustoT 
            Height          =   315
            Index           =   2
            Left            =   2925
            TabIndex        =   258
            Top             =   1785
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MOCodigo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   0
            MaxLength       =   20
            TabIndex        =   121
            Top             =   465
            Width           =   1245
         End
         Begin VB.TextBox MODescricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   3435
            MaxLength       =   50
            TabIndex        =   120
            Top             =   615
            Width           =   2235
         End
         Begin MSMask.MaskEdBox MOQuantidade 
            Height          =   315
            Index           =   2
            Left            =   5235
            TabIndex        =   117
            Top             =   195
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOHoras 
            Height          =   315
            Index           =   2
            Left            =   5790
            TabIndex        =   118
            Top             =   630
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOCusto 
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   119
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMO 
            Height          =   3075
            Index           =   2
            Left            =   75
            TabIndex        =   108
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
         Begin VB.Label MOCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   7155
            TabIndex        =   278
            Top             =   3945
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   31
            Left            =   6075
            TabIndex        =   277
            Top             =   4005
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMO 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   1
         Left            =   240
         TabIndex        =   109
         Top             =   375
         Width           =   8670
         Begin VB.TextBox MOOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   4815
            MaxLength       =   50
            TabIndex        =   303
            Top             =   2385
            Width           =   2070
         End
         Begin VB.CommandButton BotaoMOTrazer 
            Caption         =   "Trazer Calculado"
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
            Index           =   1
            Left            =   1905
            TabIndex        =   123
            Top             =   3945
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MOCustoT 
            Height          =   315
            Index           =   1
            Left            =   2805
            TabIndex        =   257
            Top             =   1965
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton BotaoMO 
            Caption         =   "Mão de Obra"
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
            Index           =   1
            Left            =   60
            TabIndex        =   122
            Top             =   3945
            Width           =   1665
         End
         Begin VB.TextBox MOCodigo 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   450
            MaxLength       =   20
            TabIndex        =   116
            Top             =   795
            Width           =   1245
         End
         Begin VB.TextBox MODescricao 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   3885
            MaxLength       =   50
            TabIndex        =   115
            Top             =   735
            Width           =   2235
         End
         Begin MSMask.MaskEdBox MOQuantidade 
            Height          =   315
            Index           =   1
            Left            =   5700
            TabIndex        =   112
            Top             =   315
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOHoras 
            Height          =   315
            Index           =   1
            Left            =   6285
            TabIndex        =   113
            Top             =   765
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOCusto 
            Height          =   315
            Index           =   1
            Left            =   3840
            TabIndex        =   114
            Top             =   105
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMO 
            Height          =   3075
            Index           =   1
            Left            =   75
            TabIndex        =   110
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
         Begin VB.Label MOCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   7140
            TabIndex        =   276
            Top             =   3975
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   30
            Left            =   6060
            TabIndex        =   275
            Top             =   4035
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMO 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   4
         Left            =   240
         TabIndex        =   230
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MOOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   2595
            MaxLength       =   50
            TabIndex        =   299
            Top             =   2205
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MOData 
            Height          =   315
            Index           =   4
            Left            =   4635
            TabIndex        =   292
            Top             =   1590
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MOCustoT 
            Height          =   315
            Index           =   4
            Left            =   2595
            TabIndex        =   260
            Top             =   1995
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MODescricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   3435
            MaxLength       =   50
            TabIndex        =   232
            Top             =   615
            Width           =   2235
         End
         Begin VB.TextBox MOCodigo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   0
            MaxLength       =   20
            TabIndex        =   231
            Top             =   465
            Width           =   1245
         End
         Begin MSMask.MaskEdBox MOQuantidade 
            Height          =   315
            Index           =   4
            Left            =   5235
            TabIndex        =   233
            Top             =   195
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOHoras 
            Height          =   315
            Index           =   4
            Left            =   5790
            TabIndex        =   234
            Top             =   630
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MOCusto 
            Height          =   315
            Index           =   4
            Left            =   3360
            TabIndex        =   235
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMO 
            Height          =   3075
            Index           =   4
            Left            =   75
            TabIndex        =   236
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
         Begin VB.Label MOCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   7170
            TabIndex        =   280
            Top             =   3975
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   32
            Left            =   6090
            TabIndex        =   279
            Top             =   4035
            Width           =   1050
         End
      End
      Begin MSComctlLib.TabStrip TabStripMO 
         Height          =   4710
         Left            =   120
         TabIndex        =   111
         Top             =   60
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   8308
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsto - Informado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsto - Calculado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Real - Informado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Real - Calculado"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Maquinas"
      Height          =   4875
      Index           =   7
      Left            =   120
      TabIndex        =   74
      Top             =   1200
      Visible         =   0   'False
      Width           =   9285
      Begin VB.Frame FrameMaq 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   3
         Left            =   240
         TabIndex        =   237
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MaqOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   1305
            MaxLength       =   50
            TabIndex        =   296
            Top             =   1140
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MaqData 
            Height          =   315
            Index           =   3
            Left            =   4710
            TabIndex        =   295
            Top             =   1470
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoMaqTrazer 
            Caption         =   "Trazer Calculado"
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
            Index           =   3
            Left            =   1920
            TabIndex        =   239
            Top             =   3945
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MaqCustoT 
            Height          =   315
            Index           =   3
            Left            =   2415
            TabIndex        =   263
            Top             =   2010
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MaqCodigo 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   0
            MaxLength       =   20
            TabIndex        =   241
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox MaqDescricao 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   975
            MaxLength       =   50
            TabIndex        =   240
            Top             =   465
            Width           =   2475
         End
         Begin VB.CommandButton BotaoMaq 
            Caption         =   "Máquinas"
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
            Index           =   3
            Left            =   45
            TabIndex        =   238
            Top             =   3945
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MaqCusto 
            Height          =   315
            Index           =   3
            Left            =   5100
            TabIndex        =   242
            Top             =   105
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqHoras 
            Height          =   315
            Index           =   3
            Left            =   5715
            TabIndex        =   243
            Top             =   480
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqQuantidade 
            Height          =   315
            Index           =   3
            Left            =   4230
            TabIndex        =   244
            Top             =   450
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaq 
            Height          =   3075
            Index           =   3
            Left            =   75
            TabIndex        =   245
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
         Begin VB.Label MaqCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   7185
            TabIndex        =   282
            Top             =   3975
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   33
            Left            =   6105
            TabIndex        =   281
            Top             =   4035
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMaq 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   2
         Left            =   240
         TabIndex        =   83
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MaqOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   3030
            MaxLength       =   50
            TabIndex        =   306
            Top             =   2535
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MaqCustoT 
            Height          =   315
            Index           =   2
            Left            =   4110
            TabIndex        =   262
            Top             =   1740
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MaqCodigo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   0
            MaxLength       =   20
            TabIndex        =   86
            Top             =   210
            Width           =   1335
         End
         Begin VB.TextBox MaqDescricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   975
            MaxLength       =   50
            TabIndex        =   85
            Top             =   675
            Width           =   2475
         End
         Begin MSMask.MaskEdBox MaqCusto 
            Height          =   315
            Index           =   2
            Left            =   5100
            TabIndex        =   87
            Top             =   315
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqHoras 
            Height          =   315
            Index           =   2
            Left            =   5715
            TabIndex        =   88
            Top             =   690
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqQuantidade 
            Height          =   315
            Index           =   2
            Left            =   4230
            TabIndex        =   89
            Top             =   660
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaq 
            Height          =   3075
            Index           =   2
            Left            =   75
            TabIndex        =   84
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
         Begin VB.Label MaqCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   7155
            TabIndex        =   286
            Top             =   3945
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   35
            Left            =   6075
            TabIndex        =   285
            Top             =   4005
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMaq 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   1
         Left            =   240
         TabIndex        =   76
         Top             =   375
         Width           =   8670
         Begin VB.TextBox MaqOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   3840
            MaxLength       =   50
            TabIndex        =   305
            Top             =   2565
            Width           =   2070
         End
         Begin VB.CommandButton BotaoMaqTrazer 
            Caption         =   "Trazer Calculado"
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
            Index           =   1
            Left            =   1845
            TabIndex        =   125
            Top             =   3945
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MaqCustoT 
            Height          =   315
            Index           =   1
            Left            =   5160
            TabIndex        =   261
            Top             =   1800
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton BotaoMaq 
            Caption         =   "Máquinas"
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
            Index           =   1
            Left            =   45
            TabIndex        =   124
            Top             =   3945
            Width           =   1665
         End
         Begin VB.TextBox MaqDescricao 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   975
            MaxLength       =   50
            TabIndex        =   79
            Top             =   465
            Width           =   2475
         End
         Begin VB.TextBox MaqCodigo 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   2700
            MaxLength       =   20
            TabIndex        =   78
            Top             =   1905
            Width           =   1335
         End
         Begin MSMask.MaskEdBox MaqCusto 
            Height          =   315
            Index           =   1
            Left            =   5100
            TabIndex        =   80
            Top             =   105
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqHoras 
            Height          =   315
            Index           =   1
            Left            =   5715
            TabIndex        =   81
            Top             =   480
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqQuantidade 
            Height          =   315
            Index           =   1
            Left            =   4230
            TabIndex        =   82
            Top             =   465
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaq 
            Height          =   3075
            Index           =   1
            Left            =   75
            TabIndex        =   77
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
         Begin VB.Label MaqCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   7155
            TabIndex        =   284
            Top             =   3960
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   34
            Left            =   6075
            TabIndex        =   283
            Top             =   4020
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMaq 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   4
         Left            =   240
         TabIndex        =   246
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MaqOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   1635
            MaxLength       =   50
            TabIndex        =   300
            Top             =   1740
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MaqData 
            Height          =   315
            Index           =   4
            Left            =   4815
            TabIndex        =   297
            Top             =   1905
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaqCustoT 
            Height          =   315
            Index           =   4
            Left            =   3765
            TabIndex        =   264
            Top             =   2490
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MaqDescricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   975
            MaxLength       =   50
            TabIndex        =   248
            Top             =   675
            Width           =   2475
         End
         Begin VB.TextBox MaqCodigo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   0
            MaxLength       =   20
            TabIndex        =   247
            Top             =   210
            Width           =   1335
         End
         Begin MSMask.MaskEdBox MaqCusto 
            Height          =   315
            Index           =   4
            Left            =   5100
            TabIndex        =   249
            Top             =   315
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqHoras 
            Height          =   315
            Index           =   4
            Left            =   5715
            TabIndex        =   250
            Top             =   690
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaqQuantidade 
            Height          =   315
            Index           =   4
            Left            =   4230
            TabIndex        =   251
            Top             =   660
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaq 
            Height          =   3075
            Index           =   4
            Left            =   75
            TabIndex        =   252
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
            Rows            =   15
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            AllowUserResizing=   1
         End
         Begin VB.Label MaqCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   7170
            TabIndex        =   288
            Top             =   3960
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   36
            Left            =   6090
            TabIndex        =   287
            Top             =   4020
            Width           =   1050
         End
      End
      Begin MSComctlLib.TabStrip TabStripMaq 
         Height          =   4710
         Left            =   120
         TabIndex        =   75
         Top             =   60
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   8308
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsto - Informado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsto - Calculado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Real - Informado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Real - Calculado"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Materiais"
      Height          =   4875
      Index           =   5
      Left            =   135
      TabIndex        =   90
      Top             =   1200
      Visible         =   0   'False
      Width           =   9285
      Begin VB.Frame FrameMP 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   2
         Left            =   240
         TabIndex        =   99
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MPOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   302
            Top             =   2055
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MPCustoT 
            Height          =   315
            Index           =   2
            Left            =   3255
            TabIndex        =   254
            Top             =   2115
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MPDescricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   1035
            MaxLength       =   50
            TabIndex        =   102
            Top             =   1320
            Width           =   2070
         End
         Begin VB.ComboBox MPUM 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "EtapaPRJ.ctx":09F8
            Left            =   4410
            List            =   "EtapaPRJ.ctx":09FA
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   1320
            Width           =   870
         End
         Begin MSMask.MaskEdBox MPProduto 
            Height          =   315
            Index           =   2
            Left            =   300
            TabIndex        =   103
            Top             =   885
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MPQuantidade 
            Height          =   315
            Index           =   2
            Left            =   2805
            TabIndex        =   104
            Top             =   720
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.0#"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MPCusto 
            Height          =   315
            Index           =   2
            Left            =   5085
            TabIndex        =   105
            Top             =   855
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMP 
            Height          =   3075
            Index           =   2
            Left            =   75
            TabIndex        =   100
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
         End
         Begin VB.Label MPCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   7125
            TabIndex        =   268
            Top             =   3930
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   20
            Left            =   6045
            TabIndex        =   267
            Top             =   3990
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMP 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   1
         Left            =   240
         TabIndex        =   92
         Top             =   375
         Width           =   8670
         Begin VB.TextBox MPOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   3420
            MaxLength       =   50
            TabIndex        =   301
            Top             =   1800
            Width           =   2070
         End
         Begin VB.CommandButton BotaoMPTrazer 
            Caption         =   "Trazer Calculado"
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
            Index           =   1
            Left            =   1860
            TabIndex        =   127
            Top             =   3930
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MPCustoT 
            Height          =   315
            Index           =   1
            Left            =   285
            TabIndex        =   253
            Top             =   1410
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton BotaoMP 
            Caption         =   "Materiais"
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
            Index           =   1
            Left            =   90
            TabIndex        =   126
            Top             =   3930
            Width           =   1665
         End
         Begin VB.ComboBox MPUM 
            Height          =   315
            Index           =   1
            Left            =   3285
            TabIndex        =   95
            Top             =   900
            Width           =   870
         End
         Begin VB.TextBox MPDescricao 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   735
            MaxLength       =   50
            TabIndex        =   94
            Top             =   990
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MPProduto 
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   96
            Top             =   495
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MPQuantidade 
            Height          =   315
            Index           =   1
            Left            =   2520
            TabIndex        =   97
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   15
            Format          =   "#,##0.0#"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MPCusto 
            Height          =   315
            Index           =   1
            Left            =   4800
            TabIndex        =   98
            Top             =   435
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMP 
            Height          =   3075
            Index           =   1
            Left            =   75
            TabIndex        =   93
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
         End
         Begin VB.Label MPCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   7125
            TabIndex        =   266
            Top             =   3915
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Left            =   6045
            TabIndex        =   265
            Top             =   3975
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMP 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   4
         Left            =   240
         TabIndex        =   214
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MPOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   3855
            MaxLength       =   50
            TabIndex        =   298
            Top             =   2685
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MPData 
            Height          =   315
            Index           =   4
            Left            =   5565
            TabIndex        =   291
            Top             =   1815
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MPCustoT 
            Height          =   315
            Index           =   4
            Left            =   2700
            TabIndex        =   256
            Top             =   2070
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox MPUM 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            ItemData        =   "EtapaPRJ.ctx":09FC
            Left            =   4410
            List            =   "EtapaPRJ.ctx":09FE
            Style           =   2  'Dropdown List
            TabIndex        =   216
            Top             =   1320
            Width           =   870
         End
         Begin VB.TextBox MPDescricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   1035
            MaxLength       =   50
            TabIndex        =   215
            Top             =   1320
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MPProduto 
            Height          =   315
            Index           =   4
            Left            =   300
            TabIndex        =   217
            Top             =   885
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MPQuantidade 
            Height          =   315
            Index           =   4
            Left            =   2805
            TabIndex        =   218
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.0#"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MPCusto 
            Height          =   315
            Index           =   4
            Left            =   5085
            TabIndex        =   219
            Top             =   855
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMP 
            Height          =   3075
            Index           =   4
            Left            =   75
            TabIndex        =   220
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
         End
         Begin VB.Label MPCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   7125
            TabIndex        =   272
            Top             =   3960
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   26
            Left            =   6045
            TabIndex        =   271
            Top             =   4020
            Width           =   1050
         End
      End
      Begin VB.Frame FrameMP 
         BorderStyle     =   0  'None
         Caption         =   "Informados"
         Height          =   4320
         Index           =   3
         Left            =   240
         TabIndex        =   205
         Top             =   375
         Visible         =   0   'False
         Width           =   8670
         Begin VB.TextBox MPOBS 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   2310
            MaxLength       =   50
            TabIndex        =   290
            Top             =   1560
            Width           =   2070
         End
         Begin MSMask.MaskEdBox MPData 
            Height          =   315
            Index           =   3
            Left            =   5310
            TabIndex        =   289
            Top             =   2025
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoMPTrazer 
            Caption         =   "Trazer Calculado"
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
            Index           =   3
            Left            =   1950
            TabIndex        =   207
            Top             =   3885
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MPCustoT 
            Height          =   315
            Index           =   3
            Left            =   3090
            TabIndex        =   255
            Top             =   2475
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MPDescricao 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   735
            MaxLength       =   50
            TabIndex        =   209
            Top             =   990
            Width           =   2070
         End
         Begin VB.ComboBox MPUM 
            Height          =   315
            Index           =   3
            Left            =   3285
            TabIndex        =   208
            Top             =   900
            Width           =   870
         End
         Begin VB.CommandButton BotaoMP 
            Caption         =   "Materiais"
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
            Index           =   3
            Left            =   90
            TabIndex        =   206
            Top             =   3900
            Width           =   1665
         End
         Begin MSMask.MaskEdBox MPProduto 
            Height          =   315
            Index           =   3
            Left            =   0
            TabIndex        =   210
            Top             =   495
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MPQuantidade 
            Height          =   315
            Index           =   3
            Left            =   2520
            TabIndex        =   211
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   15
            Format          =   "#,##0.0#"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MPCusto 
            Height          =   315
            Index           =   3
            Left            =   4800
            TabIndex        =   212
            Top             =   435
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMP 
            Height          =   3075
            Index           =   3
            Left            =   75
            TabIndex        =   213
            Top             =   150
            Width           =   8460
            _ExtentX        =   14923
            _ExtentY        =   5424
            _Version        =   393216
         End
         Begin VB.Label MPCustoTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   7125
            TabIndex        =   270
            Top             =   3870
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Custo Total:"
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
            Index           =   24
            Left            =   6045
            TabIndex        =   269
            Top             =   3930
            Width           =   1050
         End
      End
      Begin MSComctlLib.TabStrip TabStripMP 
         Height          =   4710
         Left            =   120
         TabIndex        =   91
         Top             =   45
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   8308
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsto - Informado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Previsto - Calculado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Real - Informado"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Real - Calculado"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Itens Produzidos"
      Height          =   4875
      Index           =   4
      Left            =   130
      TabIndex        =   139
      Top             =   1200
      Visible         =   0   'False
      Width           =   9285
      Begin MSMask.MaskEdBox PAQuantidade 
         Height          =   315
         Left            =   5625
         TabIndex        =   180
         Top             =   3645
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PAProduto 
         Height          =   315
         Left            =   240
         TabIndex        =   179
         Top             =   1725
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   45
         TabIndex        =   60
         ToolTipText     =   "Abre o Browse de Produtos"
         Top             =   4425
         Width           =   1200
      End
      Begin VB.CommandButton BotaoRoteiros 
         Caption         =   "Roteiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   2640
         TabIndex        =   62
         ToolTipText     =   "Visualiza o Roteiro de Fabricação"
         Top             =   4425
         Width           =   1200
      End
      Begin VB.CommandButton BotaoKit 
         Caption         =   "Kit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   1365
         TabIndex        =   61
         ToolTipText     =   "Visualiza o Kit do Produto"
         Top             =   4425
         Width           =   1200
      End
      Begin VB.TextBox PADescricao 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   144
         Top             =   1770
         Width           =   3570
      End
      Begin VB.TextBox PAObservacao 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   6240
         MaxLength       =   255
         TabIndex        =   143
         Top             =   1755
         Width           =   2475
      End
      Begin VB.ComboBox PAVersao 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "EtapaPRJ.ctx":0A00
         Left            =   4260
         List            =   "EtapaPRJ.ctx":0A02
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   2670
         Width           =   930
      End
      Begin VB.ComboBox PAUM 
         Height          =   315
         Left            =   5580
         TabIndex        =   141
         Top             =   2670
         Width           =   705
      End
      Begin MSFlexGridLib.MSFlexGrid GridPA 
         Height          =   4185
         Left            =   105
         TabIndex        =   140
         Top             =   150
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   7382
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Predecessores"
      Height          =   4875
      Index           =   3
      Left            =   135
      TabIndex        =   175
      Top             =   1200
      Visible         =   0   'False
      Width           =   9285
      Begin VB.Frame Frame4 
         Caption         =   "Predecessores"
         Height          =   4665
         Left            =   345
         TabIndex        =   176
         Top             =   120
         Width           =   8715
         Begin MSMask.MaskEdBox PredDataFim 
            Height          =   315
            Left            =   3720
            TabIndex        =   204
            Top             =   1380
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PredDataIni 
            Height          =   315
            Left            =   5040
            TabIndex        =   202
            Top             =   2040
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PredDescricao 
            Height          =   315
            Left            =   1770
            TabIndex        =   203
            Top             =   2055
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.ComboBox PredCodigo 
            Height          =   315
            Left            =   2505
            TabIndex        =   177
            Text            =   "Combo1"
            Top             =   675
            Width           =   2070
         End
         Begin MSFlexGridLib.MSFlexGrid GridPred 
            Height          =   4125
            Left            =   435
            TabIndex        =   178
            Top             =   330
            Width           =   7920
            _ExtentX        =   13970
            _ExtentY        =   7276
            _Version        =   393216
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4875
      Index           =   2
      Left            =   135
      TabIndex        =   138
      Top             =   1200
      Visible         =   0   'False
      Width           =   9285
      Begin VB.Frame Frame2 
         Caption         =   "Previsão"
         Height          =   2025
         Index           =   0
         Left            =   165
         TabIndex        =   163
         Top             =   255
         Width           =   8730
         Begin VB.Frame Frame10 
            Caption         =   "Informado"
            Height          =   810
            Left            =   135
            TabIndex        =   171
            Top             =   240
            Width           =   8490
            Begin MSMask.MaskEdBox Intervalo 
               Height          =   315
               Left            =   1080
               TabIndex        =   16
               Top             =   285
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   3
               Format          =   "###"
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataInicio 
               Height          =   300
               Left            =   5280
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   315
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataInicioInf 
               Height          =   315
               Left            =   4125
               TabIndex        =   17
               Top             =   315
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataFim 
               Height          =   300
               Left            =   7800
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   315
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataFimInf 
               Height          =   315
               Left            =   6645
               TabIndex        =   19
               Top             =   315
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "dias"
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
               Index           =   0
               Left            =   1875
               TabIndex        =   181
               Top             =   345
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Duração:"
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
               Index           =   3
               Left            =   180
               TabIndex        =   174
               Top             =   345
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Início:"
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
               Index           =   2
               Left            =   2985
               TabIndex        =   173
               Top             =   360
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Fim:"
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
               Index           =   1
               Left            =   5685
               TabIndex        =   172
               Top             =   360
               Width           =   825
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Calculado"
            Height          =   810
            Left            =   135
            TabIndex        =   164
            Top             =   1125
            Width           =   8490
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Duração:"
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
               Index           =   17
               Left            =   195
               TabIndex        =   170
               Top             =   345
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Início:"
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
               Index           =   18
               Left            =   3000
               TabIndex        =   169
               Top             =   360
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Fim:"
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
               Index           =   19
               Left            =   5700
               TabIndex        =   168
               Top             =   360
               Width           =   825
            End
            Begin VB.Label Duracao 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1065
               TabIndex        =   167
               Top             =   315
               Width           =   1665
            End
            Begin VB.Label DataInicioCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4125
               TabIndex        =   166
               Top             =   315
               Width           =   1170
            End
            Begin VB.Label DataFimCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6645
               TabIndex        =   165
               Top             =   315
               Width           =   1170
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados Reais"
         Height          =   2025
         Left            =   165
         TabIndex        =   151
         Top             =   2595
         Width           =   8730
         Begin VB.Frame Frame6 
            Caption         =   "Informado"
            Height          =   780
            Left            =   120
            TabIndex        =   159
            Top             =   270
            Width           =   8520
            Begin MSComCtl2.UpDown UpDownDataInicioReal 
               Height          =   300
               Left            =   5550
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataInicioRealInf 
               Height          =   315
               Left            =   4395
               TabIndex        =   22
               Top             =   300
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataFimReal 
               Height          =   300
               Left            =   8100
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataFimRealInf 
               Height          =   315
               Left            =   6945
               TabIndex        =   24
               Top             =   300
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PercCompRealInf 
               Height          =   315
               Left            =   2310
               TabIndex        =   21
               Top             =   285
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Fim:"
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
               Index           =   12
               Left            =   5985
               TabIndex        =   162
               Top             =   345
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Início:"
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
               Index           =   10
               Left            =   3255
               TabIndex        =   161
               Top             =   345
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Percentual Completado:"
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
               Index           =   7
               Left            =   225
               TabIndex        =   160
               Top             =   330
               Width           =   2040
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Calculado"
            Height          =   780
            Index           =   1
            Left            =   135
            TabIndex        =   152
            Top             =   1140
            Width           =   8520
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Fim:"
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
               Index           =   14
               Left            =   5940
               TabIndex        =   158
               Top             =   345
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Data Início:"
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
               Index           =   15
               Left            =   3210
               TabIndex        =   157
               Top             =   345
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Percentual Completado:"
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
               Index           =   16
               Left            =   180
               TabIndex        =   156
               Top             =   330
               Width           =   2040
            End
            Begin VB.Label PercCompRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2265
               TabIndex        =   155
               Top             =   300
               Width           =   870
            End
            Begin VB.Label DataInicioRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4380
               TabIndex        =   154
               Top             =   285
               Width           =   1170
            End
            Begin VB.Label DataFimRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6915
               TabIndex        =   153
               Top             =   285
               Width           =   1170
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Escopo"
      Height          =   4875
      Index           =   8
      Left            =   135
      TabIndex        =   128
      Top             =   1200
      Visible         =   0   'False
      Width           =   9285
      Begin VB.TextBox EscDescricao 
         Height          =   675
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   105
         Width           =   6585
      End
      Begin VB.TextBox EscExpectativa 
         Height          =   675
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   900
         Width           =   6585
      End
      Begin VB.TextBox EscFatores 
         Height          =   675
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1680
         Width           =   6585
      End
      Begin VB.TextBox EscRestricoes 
         Height          =   675
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   2460
         Width           =   6585
      End
      Begin VB.TextBox EscPremissas 
         Height          =   675
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   3240
         Width           =   6585
      End
      Begin VB.TextBox EscExclusoes 
         Height          =   675
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   4020
         Width           =   6585
      End
      Begin VB.Label Label1 
         Caption         =   "Descrição da Etapa:"
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
         Index           =   21
         Left            =   525
         TabIndex        =   134
         Top             =   105
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Expectativa do Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   22
         Left            =   285
         TabIndex        =   133
         Top             =   885
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Fatores de Sucesso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   23
         Left            =   480
         TabIndex        =   132
         Top             =   1650
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Restrições:"
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
         Index           =   25
         Left            =   1245
         TabIndex        =   131
         Top             =   2415
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Premissas:"
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
         Index           =   27
         Left            =   1290
         TabIndex        =   130
         Top             =   3180
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Exclusões específicas:"
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
         Index           =   28
         Left            =   225
         TabIndex        =   129
         Top             =   3975
         Width           =   2070
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7260
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   64
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "EtapaPRJ.ctx":0A04
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "EtapaPRJ.ctx":0B82
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "EtapaPRJ.ctx":10B4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "EtapaPRJ.ctx":123E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5505
      Left            =   105
      TabIndex        =   63
      Top             =   615
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   9710
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Predecessores"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens Prod."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Materiais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mão de Obra"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Máquinas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Escopo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Customizados"
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
Attribute VB_Name = "EtapaPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Const NUM_GRID_ARRAY = 4

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoPRJ As AdmEvento
Attribute objEventoPRJ.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoPA As AdmEvento
Attribute objEventoPA.VB_VarHelpID = -1
Private WithEvents objEventoKit As AdmEvento
Attribute objEventoKit.VB_VarHelpID = -1
Private WithEvents objEventoRoteiro As AdmEvento
Attribute objEventoRoteiro.VB_VarHelpID = -1
Private WithEvents objEventoMaq As AdmEvento
Attribute objEventoMaq.VB_VarHelpID = -1
Private WithEvents objEventoMO As AdmEvento
Attribute objEventoMO.VB_VarHelpID = -1
Private WithEvents objEventoMP As AdmEvento
Attribute objEventoMP.VB_VarHelpID = -1

Dim gobjTelaEscopo As ClassTelaEscopo
Dim gobjTelaCamposCust As ClassTelaDadosCust

Dim glNumIntPRJEtapa As Integer
Dim sProjetoAnt As String
Dim sNomeProjetoAnt As String

Dim iFrameAtual As Integer
Dim iFrameMO As Integer
Dim iFrameMaq As Integer
Dim iFrameMP As Integer

Public iAlterado As Integer

'Grid de predecessores
Dim objGridPred As AdmGrid
Dim iGrid_PredCodigo_Col As Integer
Dim iGrid_PredDescricao_Col As Integer
Dim iGrid_PredDataIni_Col As Integer
Dim iGrid_PredDataFim_Col As Integer

'Grid de Produtos acabados
Dim objGridPA As AdmGrid
Dim iGrid_PAProduto_Col As Integer
Dim iGrid_PADescricao_Col As Integer
Dim iGrid_PAObservacao_Col As Integer
Dim iGrid_PAVersao_Col As Integer
Dim iGrid_PAUM_Col As Integer
Dim iGrid_PAQuantidade_Col As Integer

'Grid de Matéria prima
Dim objGridMP(1 To NUM_GRID_ARRAY) As AdmGrid
Dim iGrid_MPProduto_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MPDescricao_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MPCusto_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MPVersao_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MPUM_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MPQuantidade_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MPCustoT_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MPData_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MPOBS_Col(1 To NUM_GRID_ARRAY) As Integer

'Grid de máquinas
Dim objGridMaq(1 To NUM_GRID_ARRAY) As AdmGrid
Dim iGrid_MaqCodigo_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MaqDescricao_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MaqCusto_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MaqHoras_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MaqQuantidade_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MaqCustoT_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MaqData_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MaqOBS_Col(1 To NUM_GRID_ARRAY) As Integer

'Grid de mão de obra
Dim objGridMO(1 To NUM_GRID_ARRAY) As AdmGrid
Dim iGrid_MOCodigo_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MODescricao_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MOCusto_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MOHoras_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MOQuantidade_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MOCustoT_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MOData_Col(1 To NUM_GRID_ARRAY) As Integer
Dim iGrid_MOOBS_Col(1 To NUM_GRID_ARRAY) As Integer

Dim iIndiceInfCalcAtual As Integer

Const FRAME_INICIAL = 1
Const FRAME_DATAS = 2
Const FRAME_Pred = 3
Const FRAME_PA = 4
Const FRAME_MP = 5
Const FRAME_MO = 6
Const FRAME_Maq = 7
Const FRAME_ESCOPO = 8
Const FRAME_DADOSCUST = 9

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Etapas do Projeto"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "EtapaPRJ"
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

Private Sub TabStripMaq_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStripMaq.SelectedItem.Index <> iFrameMaq Then

        If TabStrip_PodeTrocarTab(iFrameMaq, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameMaq(TabStripMaq.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameMaq(iFrameMaq).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameMaq = TabStripMaq.SelectedItem.Index
                
    End If
    
    iIndiceInfCalcAtual = iFrameMaq
    
End Sub

Private Sub TabStripMO_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStripMO.SelectedItem.Index <> iFrameMO Then

        If TabStrip_PodeTrocarTab(iFrameMO, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameMO(TabStripMO.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameMO(iFrameMO).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameMO = TabStripMO.SelectedItem.Index
                
    End If

    iIndiceInfCalcAtual = iFrameMO

End Sub

Private Sub TabStripMP_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStripMP.SelectedItem.Index <> iFrameMP Then

        If TabStrip_PodeTrocarTab(iFrameMP, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameMP(TabStripMP.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameMP(iFrameMP).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameMP = TabStripMP.SelectedItem.Index
                
    End If

    iIndiceInfCalcAtual = iFrameMP

End Sub

Private Sub UserControl_Initialize()
    Set gobjTelaEscopo = New ClassTelaEscopo
    Set gobjTelaCamposCust = New ClassTelaDadosCust
    Set gobjTelaEscopo.objUserControl = Me
    Set gobjTelaCamposCust.objUserControl = Me
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is NomeReduzidoPRJ Then
            Call LabelNomeRedPRJ_Click
        ElseIf Me.ActiveControl Is NomeReduzido Then
            Call LabelNomeReduzido_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        ElseIf Me.ActiveControl Is Referencia Then
            Call LabelReferencia_Click
        ElseIf Me.ActiveControl Is PAProduto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is MPProduto(iIndiceInfCalcAtual) Then
            Call BotaoMP_Click(iIndiceInfCalcAtual)
        ElseIf Me.ActiveControl Is MOCodigo(iIndiceInfCalcAtual) Then
            Call BotaoMO_Click(iIndiceInfCalcAtual)
        ElseIf Me.ActiveControl Is MaqCodigo(iIndiceInfCalcAtual) Then
            Call BotaoMaq_Click(iIndiceInfCalcAtual)
        End If
        
    End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_UnLoad

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    Set objEventoMaq = Nothing
    Set objEventoMO = Nothing
    Set objEventoMP = Nothing
    Set objEventoPRJ = Nothing
    Set objEventoPA = Nothing
    Set objEventoKit = Nothing
    Set objEventoRoteiro = Nothing
    Set gobjTelaEscopo = Nothing
    Set gobjTelaCamposCust = Nothing
    
    Set objGridPA = Nothing
    Set objGridPred = Nothing

    For iIndice = 1 To NUM_GRID_ARRAY
        Set objGridMP(iIndice) = Nothing
        Set objGridMO(iIndice) = Nothing
        Set objGridMaq(iIndice) = Nothing
    Next

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185081)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoMaq = New AdmEvento
    Set objEventoMO = New AdmEvento
    Set objEventoMP = New AdmEvento
    Set objEventoPRJ = New AdmEvento
    Set objEventoPA = New AdmEvento
    Set objEventoKit = New AdmEvento
    Set objEventoRoteiro = New AdmEvento
    
    Set objGridPred = New AdmGrid
    Set objGridPA = New AdmGrid
    
    lErro = Inicializa_GridPA(objGridPA)
    If lErro <> SUCESSO Then gError 185082
    
    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", PAProduto)
    If lErro <> SUCESSO Then gError 185083
    
    lErro = Inicializa_GridPred(objGridPred)
    If lErro <> SUCESSO Then gError 185084
   
    For iIndice = 1 To NUM_GRID_ARRAY
    
        iIndiceInfCalcAtual = iIndice
        
        Set objGridMP(iIndice) = New AdmGrid

        lErro = Inicializa_GridMP(objGridMP(iIndice))
        If lErro <> SUCESSO Then gError 185085
        
        'Inicializa a Máscara de Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", MPProduto(iIndice))
        If lErro <> SUCESSO Then gError 185086

        Set objGridMO(iIndice) = New AdmGrid

        lErro = Inicializa_GridMO(objGridMO(iIndice))
        If lErro <> SUCESSO Then gError 185087
        
        Set objGridMaq(iIndice) = New AdmGrid
        
        lErro = Inicializa_GridMaq(objGridMaq(iIndice))
        If lErro <> SUCESSO Then gError 185088
    Next
    
    lErro = Inicializa_Mascara_Projeto(Projeto)
    If lErro <> SUCESSO Then gError 189052
    
    lErro = Inicializa_Mascara_RefEtapa(Referencia)
    If lErro <> SUCESSO Then gError 189053
    
    iIndiceInfCalcAtual = INDICE_INF_PREV
        
    Call gobjTelaCamposCust.Exibe_Campos_Customizados
    
    lErro = Carrega_Usuarios(Responsavel)
    If lErro <> SUCESSO Then gError 189053
    
    If gobjFAT.iPRJExibeVistorias = MARCADO Then FrameVitorias.Visible = True
    
    iAlterado = 0
    
    iFrameAtual = 1
    iFrameMO = 1
    iFrameMP = 1
    iFrameMaq = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 185082 To 185088, 189052, 189053

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185089)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objEtapa Is Nothing) Then

        lErro = Traz_Etapa_Tela(objEtapa)
        If lErro <> SUCESSO Then gError 185090

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 185090

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185091)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objProjeto As New ClassProjetos
Dim objEtapaAux As New ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer
Dim sReferencia As String
Dim iReferenciaPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189069

    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa
    
    'Le
    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185515
    
    'Se não encontrou => Erro
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185516
    
    objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
    objEtapa.sCodigo = Codigo.Caption
    objEtapa.sNomeReduzido = NomeReduzido.Text
    objEtapa.sDescricao = Descricao.Text
    
    lErro = RefEtapa_Formata(Referencia.Text, sReferencia, iReferenciaPreenchido)
    If lErro <> SUCESSO Then gError 189097

    objEtapa.sReferencia = sReferencia
    
    objEtapaAux.sCodigo = objEtapa.sCodigo
    objEtapaAux.lNumIntDocPRJ = objEtapa.lNumIntDocPRJ

    'Lê
    lErro = CF("PRJEtapas_Le", objEtapaAux)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185517

    objEtapa.lNumIntDocEtapaPaiOrg = objEtapaAux.lNumIntDocEtapaPaiOrg
    objEtapa.iNivel = objEtapaAux.iNivel
    objEtapa.iSeq = objEtapaAux.iSeq
    objEtapa.iPosicao = objEtapaAux.iPosicao
    
    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 185092

        If lErro = SUCESSO Then
            objEtapa.lCliente = objCliente.lCodigo
        End If
            
    End If

    objEtapa.iFilialCliente = Codigo_Extrai(Filial.Text)
    
    objEtapa.sResponsavel = Responsavel.Text
    objEtapa.sObjetivo = Objetivo.Text
    objEtapa.sJustificativa = Justificativa.Text
    objEtapa.sObservacao = Observacao.Text
    
    objEtapa.dtDataInicio = StrParaDate(DataInicioInf.Text)
    objEtapa.dtDataFim = StrParaDate(DataFimInf.Text)
    objEtapa.dtDataInicioReal = StrParaDate(DataInicioRealInf.Text)
    objEtapa.dtDataFimReal = StrParaDate(DataFimRealInf.Text)
    
    objEtapa.dPercentualComplet = StrParaDbl(Val(PercCompRealInf.Text) / 100)
    
    lErro = gobjTelaEscopo.Move_Tela_Memoria(Me, objEtapa.objEscopo)
    If lErro <> SUCESSO Then gError 185093
    
    lErro = gobjTelaCamposCust.Move_Tela_Memoria(objEtapa.objCamposCust, objEtapa.objTiposCamposCust)
    If lErro <> SUCESSO Then gError 185094
    
    lErro = Move_Pred_Memoria(objEtapa)
    If lErro <> SUCESSO Then gError 185095

    lErro = Move_PA_Memoria(objEtapa)
    If lErro <> SUCESSO Then gError 185096

    lErro = Move_MP_Memoria(objEtapa)
    If lErro <> SUCESSO Then gError 185097

    lErro = Move_MO_Memoria(objEtapa)
    If lErro <> SUCESSO Then gError 185098

    lErro = Move_Maq_Memoria(objEtapa)
    If lErro <> SUCESSO Then gError 185099

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 185092 To 185099, 185515, 185517, 189069, 189097
        
        Case 185516
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185100)

    End Select

    Exit Function

End Function

Function Move_Pred_Memoria(objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objEtapaPred As ClassPRJEtapasPredecessoras
Dim objEtapaAux As ClassPRJEtapas

On Error GoTo Erro_Move_Pred_Memoria

    For iIndice = 1 To objGridPred.iLinhasExistentes
    
        Set objEtapaPred = New ClassPRJEtapasPredecessoras
        Set objEtapaAux = New ClassPRJEtapas
        
        objEtapaAux.lNumIntDocPRJ = objEtapa.lNumIntDocPRJ
        objEtapaAux.sCodigo = SCodigo_Extrai(GridPred.TextMatrix(iIndice, iGrid_PredCodigo_Col))
        
        lErro = CF("PRJEtapas_Le", objEtapaAux)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185480

        objEtapaPred.lNumIntDocEtapaPre = objEtapaAux.lNumIntDoc
        objEtapaPred.iSeq = iIndice
        
        objEtapa.colPredecessores.Add objEtapaPred
    
    Next

    Move_Pred_Memoria = SUCESSO

    Exit Function

Erro_Move_Pred_Memoria:

    Move_Pred_Memoria = gErr

    Select Case gErr
    
        Case 185480

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185118)

    End Select

    Exit Function

End Function

Function Move_PA_Memoria(objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim objEtapaPA As ClassPRJEtapaItensProd
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer

On Error GoTo Erro_Move_PA_Memoria

    For iIndice = 1 To objGridPA.iLinhasExistentes
    
        Set objEtapaPA = New ClassPRJEtapaItensProd
        
        lErro = CF("Produto_Formata", GridPA.TextMatrix(iIndice, iGrid_PAProduto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 185481
        
        objEtapaPA.sProduto = sProdutoFormatado
        objEtapaPA.dQuantidade = StrParaDbl(GridPA.TextMatrix(iIndice, iGrid_PAQuantidade_Col))
        objEtapaPA.sObservacao = GridPA.TextMatrix(iIndice, iGrid_PAObservacao_Col)
        objEtapaPA.sDescricao = GridPA.TextMatrix(iIndice, iGrid_PADescricao_Col)
        objEtapaPA.sUM = GridPA.TextMatrix(iIndice, iGrid_PAUM_Col)
        objEtapaPA.sVersao = GridPA.TextMatrix(iIndice, iGrid_PAVersao_Col)
        objEtapaPA.iSeq = iIndice
        
        objEtapa.colItensProduzidos.Add objEtapaPA
    
    Next

    Move_PA_Memoria = SUCESSO

    Exit Function

Erro_Move_PA_Memoria:

    Move_PA_Memoria = gErr

    Select Case gErr
    
        Case 185481

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185119)

    End Select

    Exit Function

End Function

Function Move_MP_Memoria(objEtapa As ClassPRJEtapas, Optional ByVal iSoTipo As Integer = 0) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim iTipo As Integer
Dim objEtapaMP As ClassPRJEtapaMateriais

On Error GoTo Erro_Move_MP_Memoria

    For iTipo = 1 To NUM_GRID_ARRAY

        If (iSoTipo = 0 And iTipo <> INDICE_CALC_REAL) Or iSoTipo = iTipo Then

            For iIndice = 1 To objGridMP(iTipo).iLinhasExistentes
            
                Set objEtapaMP = New ClassPRJEtapaMateriais
                
                lErro = CF("Produto_Formata", GridMP(iTipo).TextMatrix(iIndice, iGrid_MPProduto_Col(iTipo)), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 185483
                
                objEtapaMP.sProduto = sProdutoFormatado
                objEtapaMP.dQuantidade = StrParaDbl(GridMP(iTipo).TextMatrix(iIndice, iGrid_MPQuantidade_Col(iTipo)))
                objEtapaMP.dCusto = StrParaDbl(GridMP(iTipo).TextMatrix(iIndice, iGrid_MPCustoT_Col(iTipo)))
                objEtapaMP.sDescricao = GridMP(iTipo).TextMatrix(iIndice, iGrid_MPDescricao_Col(iTipo))
                objEtapaMP.sUM = GridMP(iTipo).TextMatrix(iIndice, iGrid_MPUM_Col(iTipo))
                objEtapaMP.iSeq = iIndice
                objEtapaMP.iTipo = iTipo
                
                If iGrid_MPData_Col(iTipo) <> 0 Then
                    objEtapaMP.dtData = StrParaDate(GridMP(iTipo).TextMatrix(iIndice, iGrid_MPData_Col(iTipo)))
                Else
                    objEtapaMP.dtData = DATA_NULA
                End If
                
                If iGrid_MPOBS_Col(iTipo) <> 0 Then
                    objEtapaMP.sObservacao = GridMP(iTipo).TextMatrix(iIndice, iGrid_MPOBS_Col(iTipo))
                End If
                
                objEtapa.colMateriaPrima.Add objEtapaMP
            
            Next
            
        End If
        
    Next
    
    Move_MP_Memoria = SUCESSO

    Exit Function

Erro_Move_MP_Memoria:

    Move_MP_Memoria = gErr

    Select Case gErr
    
        Case 185483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185120)

    End Select

    Exit Function

End Function


Function Move_MO_Memoria(objEtapa As ClassPRJEtapas, Optional ByVal iSoTipo As Integer = 0) As Long

Dim lErro As Long
Dim iTipo As Integer
Dim iIndice As Integer
Dim objEtapaMO As ClassPRJEtapaMO

On Error GoTo Erro_Move_MO_Memoria

    For iTipo = 1 To NUM_GRID_ARRAY
    
        If (iSoTipo = 0 And iTipo <> INDICE_CALC_REAL) Or iSoTipo = iTipo Then

            For iIndice = 1 To objGridMO(iTipo).iLinhasExistentes
            
                Set objEtapaMO = New ClassPRJEtapaMO
                
                objEtapaMO.iMaoDeObra = StrParaInt(GridMO(iTipo).TextMatrix(iIndice, iGrid_MOCodigo_Col(iTipo)))
                objEtapaMO.iQuantidade = StrParaInt(GridMO(iTipo).TextMatrix(iIndice, iGrid_MOQuantidade_Col(iTipo)))
                objEtapaMO.dHoras = StrParaDbl(GridMO(iTipo).TextMatrix(iIndice, iGrid_MOHoras_Col(iTipo)))
                objEtapaMO.dCusto = StrParaDbl(GridMO(iTipo).TextMatrix(iIndice, iGrid_MOCustoT_Col(iTipo)))
                objEtapaMO.sDescricao = GridMO(iTipo).TextMatrix(iIndice, iGrid_MODescricao_Col(iTipo))
                objEtapaMO.iSeq = iIndice
                objEtapaMO.iTipo = iTipo
                
                If iGrid_MOData_Col(iTipo) <> 0 Then
                    objEtapaMO.dtData = StrParaDate(GridMO(iTipo).TextMatrix(iIndice, iGrid_MOData_Col(iTipo)))
                Else
                    objEtapaMO.dtData = DATA_NULA
                End If
                
                If iGrid_MOOBS_Col(iTipo) <> 0 Then
                    objEtapaMO.sObservacao = GridMO(iTipo).TextMatrix(iIndice, iGrid_MOOBS_Col(iTipo))
                End If
                
                objEtapa.colMaoDeObra.Add objEtapaMO
            
            Next
        
        End If
        
    Next

    Move_MO_Memoria = SUCESSO

    Exit Function

Erro_Move_MO_Memoria:

    Move_MO_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185121)

    End Select

    Exit Function

End Function

Function Move_Maq_Memoria(objEtapa As ClassPRJEtapas, Optional ByVal iSoTipo As Integer = 0) As Long

Dim lErro As Long
Dim iTipo As Integer
Dim iIndice As Integer
Dim objEtapaMaq As ClassPRJEtapaMaquinas
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Move_Maq_Memoria

    For iTipo = 1 To NUM_GRID_ARRAY
    
        If (iSoTipo = 0 And iTipo <> INDICE_CALC_REAL) Or iSoTipo = iTipo Then

            For iIndice = 1 To objGridMaq(iTipo).iLinhasExistentes
            
                Set objEtapaMaq = New ClassPRJEtapaMaquinas
                Set objMaquina = New ClassMaquinas
                
                objMaquina.sNomeReduzido = GridMaq(iTipo).TextMatrix(iIndice, iGrid_MaqCodigo_Col(iTipo))
                
                lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
                If lErro <> SUCESSO And lErro <> 103100 Then gError 185484
                
                objEtapaMaq.lNumIntDocMaq = objMaquina.lNumIntDoc
                objEtapaMaq.iQuantidade = StrParaInt(GridMaq(iTipo).TextMatrix(iIndice, iGrid_MaqQuantidade_Col(iTipo)))
                objEtapaMaq.dHoras = StrParaDbl(GridMaq(iTipo).TextMatrix(iIndice, iGrid_MaqHoras_Col(iTipo)))
                objEtapaMaq.dCusto = StrParaDbl(GridMaq(iTipo).TextMatrix(iIndice, iGrid_MaqCustoT_Col(iTipo)))
                objEtapaMaq.sDescricao = GridMaq(iTipo).TextMatrix(iIndice, iGrid_MaqDescricao_Col(iTipo))
                objEtapaMaq.iSeq = iIndice
                objEtapaMaq.iTipo = iTipo
                
                If iGrid_MaqData_Col(iTipo) <> 0 Then
                    objEtapaMaq.dtData = StrParaDate(GridMaq(iTipo).TextMatrix(iIndice, iGrid_MaqData_Col(iTipo)))
                Else
                    objEtapaMaq.dtData = DATA_NULA
                End If
                
                If iGrid_MaqOBS_Col(iTipo) <> 0 Then
                    objEtapaMaq.sObservacao = GridMaq(iTipo).TextMatrix(iIndice, iGrid_MaqOBS_Col(iTipo))
                End If
                
                objEtapa.colMaquinas.Add objEtapaMaq
            
            Next
            
        End If
        
    Next

    Move_Maq_Memoria = SUCESSO

    Exit Function

Erro_Move_Maq_Memoria:

    Move_Maq_Memoria = gErr

    Select Case gErr
    
        Case 185484

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185122)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PRJEtapas"

    If Len(Trim(Projeto.ClipText)) > 0 Then

        'Lê os dados da Tela PedidoVenda
        lErro = Move_Tela_Memoria(objEtapa)
        If lErro <> SUCESSO Then gError 185101
        
    Else
        objEtapa.sCodigo = Codigo.Caption
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDocPRJ", objEtapa.lNumIntDocPRJ, 0, "NumIntDocPRJ"
    colCampoValor.Add "Codigo", objEtapa.sCodigo, STRING_ETAPAPRJ_CODIGO, "Codigo"
    'Filtros para o Sistema de Setas

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 185101

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185102)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_Tela_Preenche

    objEtapa.sCodigo = colCampoValor.Item("Codigo").vValor
    objEtapa.lNumIntDocPRJ = colCampoValor.Item("NumIntDocPRJ").vValor

    lErro = Traz_Etapa_Tela(objEtapa)
    If lErro <> SUCESSO Then gError 185103

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 185103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185104)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Projeto.ClipText)) = 0 Then gError 185105
    'If Len(Trim(Referencia.Text)) = 0 Then gError 185106
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 185107
    If Len(Trim(Codigo.Caption)) = 0 Then gError 185108

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(objEtapa)
    If lErro <> SUCESSO Then gError 185109
    
    lErro = Critica_Dados(objEtapa)
    If lErro <> SUCESSO Then gError 185110

    lErro = Trata_Alteracao(objEtapa, objEtapa.lNumIntDocPRJ, objEtapa.sCodigo)
    If lErro <> SUCESSO Then gError 185111
    
    Set objEtapa.objTela = Me

    'Grava a etapa no Banco de Dados
    lErro = CF("PRJEtapas_Grava", objEtapa)
    If lErro <> SUCESSO Then gError 185112

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 185105
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
'            Projeto.SetFocus

        Case 185106
            Call Rotina_Erro(vbOKOnly, "ERRO_REFERENCIA_ETAPA_NAO_PREENCHIDO", gErr)
'            Referencia.SetFocus

        Case 185107
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_ETAPA_NAO_PREENCHIDO", gErr)
'            NomeReduzido.SetFocus
            
        Case 185108
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ETAPA_NAO_PREENCHIDO", gErr)
            
        Case 185109, 185110, 185111, 185112

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185113)

    End Select

    Exit Function

End Function

Function Critica_Dados(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim objEtapaPA As ClassPRJEtapaItensProd
Dim objEtapaMaq As ClassPRJEtapaMaquinas
Dim objEtapaMP As ClassPRJEtapaMateriais
Dim objEtapaMO As ClassPRJEtapaMO
Dim iTipo As Integer
Dim sTipo As String
Dim iLinha As Integer

On Error GoTo Erro_Critica_Dados

    If objEtapa.dtDataInicio <> DATA_NULA And objEtapa.dtDataFim <> DATA_NULA Then
        If objEtapa.dtDataInicio > objEtapa.dtDataFim Then gError 185114
    End If

    If objEtapa.dtDataInicioReal <> DATA_NULA And objEtapa.dtDataFimReal <> DATA_NULA Then
        If objEtapa.dtDataInicioReal > objEtapa.dtDataFimReal Then gError 185115
    End If
    
    For Each objEtapaPA In objEtapa.colItensProduzidos
    
        iLinha = objEtapaPA.iSeq
    
        If objEtapaPA.dQuantidade = 0 Then gError 185486
        If Len(Trim(objEtapaPA.sDescricao)) = 0 Then gError 185487
        If Len(Trim(objEtapaPA.sVersao)) = 0 Then gError 185488
    
    Next
    
    For Each objEtapaMO In objEtapa.colMaoDeObra
    
        iTipo = objEtapaMO.iTipo
        iLinha = objEtapaMO.iSeq
        
        If objEtapaMO.dHoras = 0 Then gError 185489
        If Len(Trim(objEtapaMO.sDescricao)) = 0 Then gError 185490
    
    Next
    
    For Each objEtapaMaq In objEtapa.colMaquinas
    
        iTipo = objEtapaMaq.iTipo
        iLinha = objEtapaMaq.iSeq
    
        If objEtapaMaq.dHoras = 0 Then gError 185491
        If Len(Trim(objEtapaMaq.sDescricao)) = 0 Then gError 185492
    
    Next
    
    For Each objEtapaMP In objEtapa.colMateriaPrima
    
        iTipo = objEtapaMP.iTipo
        iLinha = objEtapaMP.iSeq
   
        If objEtapaMP.dQuantidade = 0 Then gError 185493
        If Len(Trim(objEtapaMP.sDescricao)) = 0 Then gError 185494
    
    Next

    GL_objMDIForm.MousePointer = vbDefault

    Critica_Dados = SUCESSO

    Exit Function

Erro_Critica_Dados:

    Critica_Dados = gErr

    Select Case gErr
            
        Case 185114
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
'            DataInicioInf.SetFocus

        Case 185115
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
'            DataInicioRealInf.SetFocus
            
        Case 185486
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_PAQUANTIDADE_NAO_PREECHIDA", gErr, iLinha)

        Case 185487
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_PADESCRICAO_NAO_PREECHIDA", gErr, iLinha)

        Case 185488
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_PAVERSAO_NAO_PREECHIDA", gErr, iLinha)

        Case 185489
            Call Tipo_Nome(iTipo, sTipo)
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MOHORAS_NAO_PREECHIDA", gErr, sTipo, iLinha)

        Case 185490
            Call Tipo_Nome(iTipo, sTipo)
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MODESCRICAO_NAO_PREECHIDA", gErr, sTipo, iLinha)

        Case 185491
            Call Tipo_Nome(iTipo, sTipo)
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MAQHORAS_NAO_PREECHIDA", gErr, sTipo, iLinha)

        Case 185492
            Call Tipo_Nome(iTipo, sTipo)
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MAQDESCRICAO_NAO_PREECHIDA", gErr, sTipo, iLinha)

        Case 185493
            Call Tipo_Nome(iTipo, sTipo)
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MPQUANTIDADE_NAO_PREECHIDA", gErr, sTipo, iLinha)

        Case 185494
            Call Tipo_Nome(iTipo, sTipo)
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MPDESCRICAO_NAO_PREECHIDA", gErr, sTipo, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185116)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Etapas() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_Etapas

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Filial.ListIndex = -1
    
    Call Grid_Limpa(objGridPA)
    Call Grid_Limpa(objGridPred)
    For iIndice = 1 To NUM_GRID_ARRAY
        Call Grid_Limpa(objGridMP(iIndice))
        Call Grid_Limpa(objGridMO(iIndice))
        Call Grid_Limpa(objGridMaq(iIndice))
    Next
    
    For iIndice = 1 To NUM_GRID_ARRAY
        MPCustoTotal(iIndice).Caption = ""
        MOCustoTotal(iIndice).Caption = ""
        MaqCustoTotal(iIndice).Caption = ""
    Next
    
    Responsavel.Text = ""
    RespUsu.Value = vbUnchecked
    
    Duracao.Caption = ""
    DataInicioCalc.Caption = ""
    DataFimCalc.Caption = ""
    PercCompRealCalc.Caption = ""
    DataInicioRealCalc.Caption = ""
    DataFimRealCalc.Caption = ""
    Codigo.Caption = ""
    
    DataVistoria.Caption = ""
    ValidadeVistoria.Caption = ""

    sProjetoAnt = ""
    sNomeProjetoAnt = ""
    glNumIntPRJEtapa = 0

    iAlterado = 0

    Limpa_Tela_Etapas = SUCESSO

    Exit Function

Erro_Limpa_Tela_Etapas:

    Limpa_Tela_Etapas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185117)

    End Select

    Exit Function

End Function

Function Traz_Etapa_Tela(objEtapas As ClassPRJEtapas) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim objEtapaAux As ClassPRJEtapas

On Error GoTo Erro_Traz_Etapa_Tela

    Call Limpa_Tela_Etapas

    'Lê a Etapa que está sendo Passada
    lErro = CF("PRJEtapas_Le", objEtapas)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185123

    Codigo.Caption = objEtapas.sCodigo
    NomeReduzido.Text = objEtapas.sNomeReduzido
    Referencia.Text = objEtapas.sReferencia

    If lErro = SUCESSO Then
    
        glNumIntPRJEtapa = objEtapas.lNumIntDoc
    
        objProjeto.lNumIntDoc = objEtapas.lNumIntDocPRJ
        
        'Lê o Projetos que está sendo Passado
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185124
        
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 189108
            
        Call Projeto_Validate(bSGECancelDummy)
    
        lErro = CF("PRJEtapas_Le_Projeto", objProjeto)
        If lErro <> SUCESSO Then gError 189256

        Call objProjeto.Calcula_Dados_Calculados
        
        For Each objEtapaAux In objProjeto.colEtapas
            If objEtapaAux.sCodigo = objEtapas.sCodigo Then
                objEtapas.objDadosCalc = objEtapaAux.objDadosCalc
                Exit For
            End If
        Next
    
        Codigo.Caption = objEtapas.sCodigo
        
        Descricao.Text = objEtapas.sDescricao

        If objEtapas.lCliente <> 0 Then

            Cliente.Text = CStr(objEtapas.lCliente)
            Call Cliente_Validate(bSGECancelDummy)
            
            Filial.Text = objEtapas.iFilialCliente
            Call Filial_Validate(bSGECancelDummy)
        
        End If

        Responsavel.Text = objEtapas.sResponsavel
        Call Responsavel_Validate(bSGECancelDummy)
        
        Objetivo.Text = objEtapas.sObjetivo
        Justificativa.Text = objEtapas.sJustificativa
        Observacao.Text = objEtapas.sObservacao

        If objEtapas.dtDataInicio <> DATA_NULA Then
            DataInicioInf.PromptInclude = False
            DataInicioInf.Text = Format(objEtapas.dtDataInicio, "dd/mm/yy")
            DataInicioInf.PromptInclude = True
        End If

        If objEtapas.dtDataFim <> DATA_NULA Then
            DataFimInf.PromptInclude = False
            DataFimInf.Text = Format(objEtapas.dtDataFim, "dd/mm/yy")
            DataFimInf.PromptInclude = True
        End If
        
        If objEtapas.dtDataInicio <> DATA_NULA And objEtapas.dtDataFim <> DATA_NULA Then
            Intervalo.PromptInclude = False
            Intervalo.Text = DateDiff("d", objEtapas.dtDataInicio, objEtapas.dtDataFim) + 1
            Intervalo.PromptInclude = True
        End If

        If objEtapas.dtDataInicioReal <> DATA_NULA Then
            DataInicioRealInf.PromptInclude = False
            DataInicioRealInf.Text = Format(objEtapas.dtDataInicioReal, "dd/mm/yy")
            DataInicioRealInf.PromptInclude = True
        End If

        If objEtapas.dtDataFimReal <> DATA_NULA Then
            DataFimRealInf.PromptInclude = False
            DataFimRealInf.Text = Format(objEtapas.dtDataFimReal, "dd/mm/yy")
            DataFimRealInf.PromptInclude = True
        End If

        PercCompRealInf.Text = CStr(objEtapas.dPercentualComplet * 100)
        
        If objEtapas.objDadosCalc.dtDataIniPrev <> DATA_NULA Then
            DataInicioCalc.Caption = Format(objEtapas.objDadosCalc.dtDataIniPrev, "dd/mm/yyyy")
        End If

        If objEtapas.objDadosCalc.dtDataFimPrev <> DATA_NULA Then
            DataFimCalc.Caption = Format(objEtapas.objDadosCalc.dtDataFimPrev, "dd/mm/yyyy")
        End If

        If objEtapas.objDadosCalc.dtDataIniReal <> DATA_NULA Then
            DataInicioRealCalc.Caption = Format(objEtapas.objDadosCalc.dtDataIniReal, "dd/mm/yyyy")
        End If

        If objEtapas.objDadosCalc.dtDataFimReal <> DATA_NULA Then
            DataFimRealCalc.Caption = Format(objEtapas.objDadosCalc.dtDataFimReal, "dd/mm/yyyy")
        End If
        
        If objEtapas.objDadosCalc.dtDataFimPrev <> DATA_NULA And objEtapas.objDadosCalc.dtDataIniPrev <> DATA_NULA Then
            Duracao.Caption = CStr(1 + DateDiff("d", objEtapas.objDadosCalc.dtDataIniPrev, objEtapas.objDadosCalc.dtDataFimPrev))
        End If

        PercCompRealCalc.Caption = Format(objEtapas.objDadosCalc.dPercentualComplet, "PERCENT")
        
        objEtapas.objEscopo.lNumIntDoc = objEtapas.lNumIntDocEscopo

        lErro = gobjTelaEscopo.Traz_PRJEscopo_Tela(Me, objEtapas.objEscopo)
        If lErro <> SUCESSO Then gError 185125
        
        objEtapas.objCamposCust.iTipoNumIntDocOrigem = CAMPO_CUSTOMIZADO_TIPO_ETAPA
        objEtapas.objCamposCust.lNumIntDocOrigem = objEtapas.lNumIntDoc

        lErro = gobjTelaCamposCust.Traz_CamposCustomizados_Tela(objEtapas.objCamposCust)
        If lErro <> SUCESSO Then gError 185126
        
        lErro = Traz_Pred_Tela(objEtapas)
        If lErro <> SUCESSO Then gError 185127
    
        lErro = Traz_PA_Tela(objEtapas)
        If lErro <> SUCESSO Then gError 185128
    
        lErro = Traz_MP_Tela(objEtapas)
        If lErro <> SUCESSO Then gError 185129
    
        lErro = Traz_MO_Tela(objEtapas)
        If lErro <> SUCESSO Then gError 185130
    
        lErro = Traz_Maq_Tela(objEtapas)
        If lErro <> SUCESSO Then gError 185131
        
        If objEtapas.dtDataVistoria <> DATA_NULA Then DataVistoria.Caption = Format(objEtapas.dtDataVistoria, "dd/mm/yyyy")
        If objEtapas.dtValidadeVistoria <> DATA_NULA Then ValidadeVistoria.Caption = Format(objEtapas.dtValidadeVistoria, "dd/mm/yyyy")
        
    Else
    
        glNumIntPRJEtapa = 0

    End If

    iAlterado = 0

    Traz_Etapa_Tela = SUCESSO

    Exit Function

Erro_Traz_Etapa_Tela:

    Traz_Etapa_Tela = gErr

    Select Case gErr

        Case 185123 To 185131, 189108, 189256

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185132)

    End Select

    Exit Function

End Function

Function Traz_Pred_Tela(objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim objEtapaPred As ClassPRJEtapasPredecessoras
Dim objEtapaAux As ClassPRJEtapas
Dim iLinha As Integer

On Error GoTo Erro_Traz_Pred_Tela

    'Exibe os dados da coleção de Competencias na tela
    For Each objEtapaPred In objEtapa.colPredecessores
        
        iLinha = iLinha + 1
        
        Set objEtapaAux = New ClassPRJEtapas
        
        objEtapaAux.lNumIntDoc = objEtapaPred.lNumIntDocEtapaPre
        
        lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapaAux)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185479
        
        GridPred.TextMatrix(iLinha, iGrid_PredCodigo_Col) = objEtapaAux.sCodigo & SEPARADOR & objEtapaAux.sNomeReduzido
        If objEtapaAux.dtDataInicio <> DATA_NULA Then
            GridPred.TextMatrix(iLinha, iGrid_PredDataIni_Col) = Format(objEtapaAux.dtDataInicio, "dd/mm/yyyy")
        Else
            GridPred.TextMatrix(iLinha, iGrid_PredDataIni_Col) = ""
        End If
        If objEtapaAux.dtDataFim <> DATA_NULA Then
            GridPred.TextMatrix(iLinha, iGrid_PredDataFim_Col) = Format(objEtapaAux.dtDataFim, "dd/mm/yyyy")
        Else
            GridPred.TextMatrix(iLinha, iGrid_PredDataFim_Col) = ""
        End If
        GridPred.TextMatrix(iLinha, iGrid_PredDescricao_Col) = objEtapaAux.sDescricao
    
    Next

    objGridPred.iLinhasExistentes = objEtapa.colPredecessores.Count

    Traz_Pred_Tela = SUCESSO

    Exit Function

Erro_Traz_Pred_Tela:

    Traz_Pred_Tela = gErr

    Select Case gErr
    
        Case 185479

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185133)

    End Select

    Exit Function

End Function

Function Traz_PA_Tela(objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim objEtapaPA As ClassPRJEtapaItensProd
Dim iLinha As Integer
Dim sProdutoMascarado As String

On Error GoTo Erro_Traz_PA_Tela

    'Exibe os dados da coleção de Competencias na tela
    For Each objEtapaPA In objEtapa.colItensProduzidos
        
        iLinha = iLinha + 1
                
        lErro = Mascara_RetornaProdutoTela(objEtapaPA.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 185470
        
        PAProduto.PromptInclude = False
        PAProduto.Text = sProdutoMascarado
        PAProduto.PromptInclude = True
        
        GridPA.TextMatrix(iLinha, iGrid_PAObservacao_Col) = objEtapaPA.sObservacao
        GridPA.TextMatrix(iLinha, iGrid_PADescricao_Col) = objEtapaPA.sDescricao
        GridPA.TextMatrix(iLinha, iGrid_PAProduto_Col) = PAProduto.Text
        GridPA.TextMatrix(iLinha, iGrid_PAQuantidade_Col) = Formata_Estoque(objEtapaPA.dQuantidade)
        GridPA.TextMatrix(iLinha, iGrid_PAUM_Col) = objEtapaPA.sUM
        GridPA.TextMatrix(iLinha, iGrid_PAVersao_Col) = objEtapaPA.sVersao
    
    Next

    objGridPA.iLinhasExistentes = objEtapa.colItensProduzidos.Count

    Traz_PA_Tela = SUCESSO

    Exit Function

Erro_Traz_PA_Tela:

    Traz_PA_Tela = gErr

    Select Case gErr
    
        Case 185470

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185134)

    End Select

    Exit Function

End Function

Function Traz_MP_Tela(objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim objPRJEtapaMP As ClassPRJEtapaMateriais
Dim iIndice As Integer
Dim colAux As Collection

On Error GoTo Erro_Traz_MP_Tela

    For iIndice = 1 To NUM_GRID_ARRAY
    
        Set colAux = New Collection
        
        For Each objPRJEtapaMP In objEtapa.colMateriaPrima
            If objPRJEtapaMP.iTipo = iIndice Then colAux.Add objPRJEtapaMP
        Next
        
        lErro = Traz_MP_Tela_Indice(colAux, iIndice)
        If lErro <> SUCESSO Then gError 185444
    
    Next

    Traz_MP_Tela = SUCESSO

    Exit Function

Erro_Traz_MP_Tela:

    Traz_MP_Tela = gErr

    Select Case gErr
    
        Case 185444

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185135)

    End Select

    Exit Function

End Function

Function Traz_MP_Tela_Indice(ByVal colMP As Collection, ByVal iTipo As Integer) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objPRJEtapaMP As ClassPRJEtapaMateriais
Dim sProdutoMascarado As String
Dim dCustoTotal As Double

On Error GoTo Erro_Traz_MP_Tela_Indice
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objPRJEtapaMP In colMP
        
        iLinha = iLinha + 1
                
        lErro = Mascara_RetornaProdutoTela(objPRJEtapaMP.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 185446
        
        MPProduto(iTipo).PromptInclude = False
        MPProduto(iTipo).Text = sProdutoMascarado
        MPProduto(iTipo).PromptInclude = True
        
        GridMP(iTipo).TextMatrix(iLinha, iGrid_MPCusto_Col(iTipo)) = Format(objPRJEtapaMP.dCusto / objPRJEtapaMP.dQuantidade, "STANDARD")
        GridMP(iTipo).TextMatrix(iLinha, iGrid_MPDescricao_Col(iTipo)) = objPRJEtapaMP.sDescricao
        GridMP(iTipo).TextMatrix(iLinha, iGrid_MPProduto_Col(iTipo)) = MPProduto(iTipo).Text
        GridMP(iTipo).TextMatrix(iLinha, iGrid_MPQuantidade_Col(iTipo)) = Formata_Estoque(objPRJEtapaMP.dQuantidade)
        GridMP(iTipo).TextMatrix(iLinha, iGrid_MPUM_Col(iTipo)) = objPRJEtapaMP.sUM
        GridMP(iTipo).TextMatrix(iLinha, iGrid_MPCustoT_Col(iTipo)) = Format(objPRJEtapaMP.dCusto, "STANDARD")

        If iGrid_MPData_Col(iTipo) <> 0 Then
            GridMP(iTipo).TextMatrix(iLinha, iGrid_MPData_Col(iTipo)) = Format(objPRJEtapaMP.dtData, "dd/mm/yyyy")
        End If
        
        If iGrid_MPOBS_Col(iTipo) <> 0 Then
            GridMP(iTipo).TextMatrix(iLinha, iGrid_MPOBS_Col(iTipo)) = objPRJEtapaMP.sObservacao
        End If
        
        dCustoTotal = dCustoTotal + objPRJEtapaMP.dCusto
    
    Next

    MPCustoTotal(iTipo).Caption = Format(dCustoTotal, "STANDARD")

    objGridMP(iTipo).iLinhasExistentes = colMP.Count

    Traz_MP_Tela_Indice = SUCESSO

    Exit Function

Erro_Traz_MP_Tela_Indice:

    Traz_MP_Tela_Indice = gErr

    Select Case gErr
    
        Case 185446

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185447)

    End Select

    Exit Function

End Function

Function Traz_MO_Tela(objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPRJEtapaMO As ClassPRJEtapaMO
Dim colAux As Collection

On Error GoTo Erro_Traz_MO_Tela

    For iIndice = 1 To NUM_GRID_ARRAY
    
        Set colAux = New Collection
        
        For Each objPRJEtapaMO In objEtapa.colMaoDeObra
            If objPRJEtapaMO.iTipo = iIndice Then colAux.Add objPRJEtapaMO
        Next
        
        lErro = Traz_MO_Tela_Indice(colAux, iIndice)
        If lErro <> SUCESSO Then gError 185451
    
    Next

    Traz_MO_Tela = SUCESSO

    Exit Function

Erro_Traz_MO_Tela:

    Traz_MO_Tela = gErr

    Select Case gErr
    
        Case 185451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185136)

    End Select

    Exit Function

End Function

Function Traz_MO_Tela_Indice(ByVal colMO As Collection, ByVal iTipo As Integer) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objPRJEtapaMO As ClassPRJEtapaMO
Dim dCustoTotal As Double

On Error GoTo Erro_Traz_MO_Tela_Indice
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objPRJEtapaMO In colMO
        
        iLinha = iLinha + 1
        
        GridMO(iTipo).TextMatrix(iLinha, iGrid_MOCusto_Col(iTipo)) = Format(objPRJEtapaMO.dCusto / objPRJEtapaMO.dHoras / objPRJEtapaMO.iQuantidade, "STANDARD")
        GridMO(iTipo).TextMatrix(iLinha, iGrid_MODescricao_Col(iTipo)) = objPRJEtapaMO.sDescricao
        GridMO(iTipo).TextMatrix(iLinha, iGrid_MOCodigo_Col(iTipo)) = objPRJEtapaMO.iMaoDeObra
        GridMO(iTipo).TextMatrix(iLinha, iGrid_MOQuantidade_Col(iTipo)) = objPRJEtapaMO.iQuantidade
        GridMO(iTipo).TextMatrix(iLinha, iGrid_MOHoras_Col(iTipo)) = Formata_Estoque(objPRJEtapaMO.dHoras)
        GridMO(iTipo).TextMatrix(iLinha, iGrid_MOCustoT_Col(iTipo)) = Format(objPRJEtapaMO.dCusto, "STANDARD")
        
        If iGrid_MOData_Col(iTipo) <> 0 Then
            GridMO(iTipo).TextMatrix(iLinha, iGrid_MOData_Col(iTipo)) = Format(objPRJEtapaMO.dtData, "dd/mm/yyyy")
        End If
        
        If iGrid_MOOBS_Col(iTipo) <> 0 Then
            GridMO(iTipo).TextMatrix(iLinha, iGrid_MOOBS_Col(iTipo)) = objPRJEtapaMO.sObservacao
        End If
        
        dCustoTotal = dCustoTotal + objPRJEtapaMO.dCusto
    
    Next

    MOCustoTotal(iTipo).Caption = Format(dCustoTotal, "STANDARD")

    objGridMO(iTipo).iLinhasExistentes = colMO.Count

    Traz_MO_Tela_Indice = SUCESSO

    Exit Function

Erro_Traz_MO_Tela_Indice:

    Traz_MO_Tela_Indice = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185452)

    End Select

    Exit Function

End Function

Function Traz_Maq_Tela(objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long
Dim objPRJEtapaMaq As ClassPRJEtapaMaquinas
Dim colAux As Collection
Dim iIndice As Integer

On Error GoTo Erro_Traz_Maq_Tela

    For iIndice = 1 To NUM_GRID_ARRAY
    
        Set colAux = New Collection
        
        For Each objPRJEtapaMaq In objEtapa.colMaquinas
            If objPRJEtapaMaq.iTipo = iIndice Then colAux.Add objPRJEtapaMaq
        Next
        
        lErro = Traz_Maq_Tela_Indice(colAux, iIndice)
        If lErro <> SUCESSO Then gError 185454
    
    Next
    
    Traz_Maq_Tela = SUCESSO

    Exit Function

Erro_Traz_Maq_Tela:

    Traz_Maq_Tela = gErr

    Select Case gErr
    
        Case 185454

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185137)

    End Select

    Exit Function

End Function

Function Traz_Maq_Tela_Indice(ByVal colMaq As Collection, ByVal iTipo As Integer) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objPRJEtapaMaq As ClassPRJEtapaMaquinas
Dim objMaquina As ClassMaquinas
Dim dCustoTotal As Double

On Error GoTo Erro_Traz_Maq_Tela_Indice
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objPRJEtapaMaq In colMaq
        
        iLinha = iLinha + 1
        
        Set objMaquina = New ClassMaquinas
        
        objMaquina.lNumIntDoc = objPRJEtapaMaq.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 185471
        
        GridMaq(iTipo).TextMatrix(iLinha, iGrid_MaqCusto_Col(iTipo)) = Format(objPRJEtapaMaq.dCusto / objPRJEtapaMaq.dHoras / objPRJEtapaMaq.iQuantidade, "STANDARD")
        GridMaq(iTipo).TextMatrix(iLinha, iGrid_MaqDescricao_Col(iTipo)) = objPRJEtapaMaq.sDescricao
        GridMaq(iTipo).TextMatrix(iLinha, iGrid_MaqCodigo_Col(iTipo)) = objMaquina.sNomeReduzido
        GridMaq(iTipo).TextMatrix(iLinha, iGrid_MaqQuantidade_Col(iTipo)) = objPRJEtapaMaq.iQuantidade
        GridMaq(iTipo).TextMatrix(iLinha, iGrid_MaqHoras_Col(iTipo)) = Formata_Estoque(objPRJEtapaMaq.dHoras)
        GridMaq(iTipo).TextMatrix(iLinha, iGrid_MaqCustoT_Col(iTipo)) = Format(objPRJEtapaMaq.dCusto, "STANDARD")
    
        If iGrid_MaqData_Col(iTipo) <> 0 Then
            GridMaq(iTipo).TextMatrix(iLinha, iGrid_MaqData_Col(iTipo)) = Format(objPRJEtapaMaq.dtData, "dd/mm/yyyy")
        End If
        
        If iGrid_MaqOBS_Col(iTipo) <> 0 Then
            GridMaq(iTipo).TextMatrix(iLinha, iGrid_MaqOBS_Col(iTipo)) = objPRJEtapaMaq.sObservacao
        End If
    
        dCustoTotal = dCustoTotal + objPRJEtapaMaq.dCusto
        
    Next
    
    MaqCustoTotal(iTipo).Caption = Format(dCustoTotal, "STANDARD")

    objGridMaq(iTipo).iLinhasExistentes = colMaq.Count

    Traz_Maq_Tela_Indice = SUCESSO

    Exit Function

Erro_Traz_Maq_Tela_Indice:

    Traz_Maq_Tela_Indice = gErr

    Select Case gErr
    
        Case 185471

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185453)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 185138

    'Limpa Tela
    Call Limpa_Tela_Etapas

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 185138

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185139)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185140)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 185141

    Call Limpa_Tela_Etapas

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 185141

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185142)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objEtapas As New ClassPRJEtapas
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Codigo.Caption)) = 0 Then gError 185143
    If Len(Trim(Projeto.ClipText)) = 0 Then gError 185144
    
    lErro = Move_Tela_Memoria(objEtapas)
    If lErro <> SUCESSO Then gError 185145

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PRJETAPAS", objEtapas.sCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("PRJEtapas_Exclui", objEtapas)
        If lErro <> SUCESSO Then gError 185146

        'Limpa Tela
        Call Limpa_Tela_Etapas

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
        
        Case 185143
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ETAPA_NAO_PREENCHIDO2", gErr)
        
        Case 185144
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            
        Case 185145, 185146

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185147)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate

    'Verifica se NomeReduzido está preenchida
    If Len(Trim(NomeReduzido.Text)) <> 0 Then

       '#######################################
       'CRITICA NomeReduzido
       '#######################################

    End If

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185148)

    End Select

    Exit Sub

End Sub

Private Sub NomeReduzido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Verifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

       '#######################################
       'CRITICA Descricao
       '#######################################

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185149)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 185150
                   
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 185151

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 185150, 185151

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185152)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 185153

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 185154

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 185155

        If lErro = 17660 Then

            'Lê o Cliente
            objCliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 185156

            'Se encontrou o Cliente
            If lErro = SUCESSO Then
                
                objFilialCliente.lCodCliente = objCliente.lCodigo

                gError 185157
            
            End If
            
        End If
        
        If iCodigo <> 0 Then
        
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
        Else
            
            objCliente.lCodigo = 0
            objFilialCliente.iCodFilial = 0
            
        End If
        
    'Não encontrou a STRING
    ElseIf lErro = 6731 Then
        
        'trecho incluido por Leo em 17/04/02
        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 185158
        
        If lErro = SUCESSO Then gError 185159
        
    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 185153, 185155

        Case 185154
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 185156, 185158 'tratado na rotina chamada

        Case 185157
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 185159
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185160)

    End Select

    Exit Sub


End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Responsavel_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Responsavel_Validate

    RespUsu.Value = vbUnchecked

    'Verifica se Responsavel está preenchida
    If Len(Trim(Responsavel.Text)) <> 0 Then

        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = Responsavel.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then RespUsu.Value = vbChecked

    End If

    Exit Sub

Erro_Responsavel_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185161)

    End Select

    Exit Sub

End Sub

Private Sub Responsavel_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Objetivo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Objetivo_Validate

    'Verifica se Objetivo está preenchida
    If Len(Trim(Objetivo.Text)) <> 0 Then

       '#######################################
       'CRITICA Objetivo
       '#######################################

    End If

    Exit Sub

Erro_Objetivo_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185162)

    End Select

    Exit Sub

End Sub

Private Sub Objetivo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Justificativa_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Justificativa_Validate

    'Verifica se Justificativa está preenchida
    If Len(Trim(Justificativa.Text)) <> 0 Then

       '#######################################
       'CRITICA Justificativa
       '#######################################

    End If

    Exit Sub

Erro_Justificativa_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185163)

    End Select

    Exit Sub

End Sub

Private Sub Justificativa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Observacao_Validate

    'Verifica se Observacao está preenchida
    If Len(Trim(Observacao.Text)) <> 0 Then

       '#######################################
       'CRITICA Observacao
       '#######################################

    End If

    Exit Sub

Erro_Observacao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185164)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_DownClick

    DataInicioInf.SetFocus

    If Len(DataInicioInf.ClipText) > 0 Then

        sData = DataInicioInf.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 185165

        DataInicioInf.Text = sData
        
        Call DataInicioInf_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 185165

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185166)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_UpClick

    DataInicioInf.SetFocus

    If Len(Trim(DataInicioInf.ClipText)) > 0 Then

        sData = DataInicioInf.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 185167

        DataInicioInf.Text = sData
        
        Call DataInicioInf_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 185167

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185168)

    End Select

    Exit Sub

End Sub

Private Sub DataInicioInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicioInf, iAlterado)
    
End Sub

Private Sub DataInicioInf_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData1 As Date
Dim dtData2 As Date

On Error GoTo Erro_DataInicioInf_Validate

    If Len(Trim(DataInicioInf.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicioInf.Text)
        If lErro <> SUCESSO Then gError 185169

        dtData1 = StrParaDate(DataInicioInf.Text)

        If StrParaDate(DataFimInf.Text) <> DATA_NULA Then
        
            dtData2 = StrParaDate(DataFimInf.Text)
            
            Intervalo.PromptInclude = False
            Intervalo.Text = CStr(1 + DateDiff("d", dtData1, dtData2))
            Intervalo.PromptInclude = True
        
        Else
        
            If Len(Trim(Intervalo.Text)) <> 0 Then
        
                DataFimInf.PromptInclude = False
                DataFimInf.Text = Format(DateAdd("d", StrParaInt(Intervalo.Text) - 1, dtData1), "dd/mm/yy")
                DataFimInf.PromptInclude = True
        
            End If
            
        End If
    
    End If

    Exit Sub

Erro_DataInicioInf_Validate:

    Cancel = True

    Select Case gErr

        Case 185169

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185170)

    End Select

    Exit Sub

End Sub

Private Sub DataInicioInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFim_DownClick

    DataFimInf.SetFocus

    If Len(DataFimInf.ClipText) > 0 Then

        sData = DataFimInf.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 185171

        DataFimInf.Text = sData
        
        Call DataFimInf_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataFim_DownClick:

    Select Case gErr

        Case 185171

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185172)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFim_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFim_UpClick

    DataFimInf.SetFocus

    If Len(Trim(DataFimInf.ClipText)) > 0 Then

        sData = DataFimInf.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 185173

        DataFimInf.Text = sData
        
        Call DataFimInf_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataFim_UpClick:

    Select Case gErr

        Case 185173

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185174)

    End Select

    Exit Sub

End Sub

Private Sub DataFimInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFimInf, iAlterado)
    
End Sub

Private Sub DataFimInf_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtData1 As Date
Dim dtData2 As Date

On Error GoTo Erro_DataFimInf_Validate

    If Len(Trim(DataFimInf.ClipText)) <> 0 Then

        lErro = Data_Critica(DataFimInf.Text)
        If lErro <> SUCESSO Then gError 185175

        dtData2 = StrParaDate(DataFimInf.Text)

        If StrParaDate(DataInicioInf.Text) <> DATA_NULA Then
        
            dtData1 = StrParaDate(DataInicioInf.Text)
            
            Intervalo.PromptInclude = False
            Intervalo.Text = CStr(1 + DateDiff("d", dtData1, dtData2))
            Intervalo.PromptInclude = True
        
        Else
        
            If Len(Trim(Intervalo.Text)) <> 0 Then
        
                DataInicioInf.PromptInclude = False
                DataInicioInf.Text = Format(DateAdd("d", -(StrParaInt(Intervalo.Text) - 1), dtData2), "dd/mm/yy")
                DataInicioInf.PromptInclude = True
        
            End If
            
        End If
        
    End If

    Exit Sub

Erro_DataFimInf_Validate:

    Cancel = True

    Select Case gErr

        Case 185175

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185176)

    End Select

    Exit Sub

End Sub

Private Sub DataFimInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataInicioReal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicioReal_DownClick

    DataInicioRealInf.SetFocus

    If Len(DataInicioRealInf.ClipText) > 0 Then

        sData = DataInicioRealInf.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 185177

        DataInicioRealInf.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicioReal_DownClick:

    Select Case gErr

        Case 185177

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185178)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicioReal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicioReal_UpClick

    DataInicioRealInf.SetFocus

    If Len(Trim(DataInicioRealInf.ClipText)) > 0 Then

        sData = DataInicioRealInf.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 185179

        DataInicioRealInf.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicioReal_UpClick:

    Select Case gErr

        Case 185179

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185180)

    End Select

    Exit Sub

End Sub

Private Sub DataInicioRealInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicioRealInf, iAlterado)
    
End Sub

Private Sub DataInicioRealInf_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicioRealInf_Validate

    If Len(Trim(DataInicioRealInf.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicioRealInf.Text)
        If lErro <> SUCESSO Then gError 185181

    End If

    Exit Sub

Erro_DataInicioRealInf_Validate:

    Cancel = True

    Select Case gErr

        Case 185181

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185182)

    End Select

    Exit Sub

End Sub

Private Sub DataInicioRealInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataFimReal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFimReal_DownClick

    DataFimRealInf.SetFocus

    If Len(DataFimRealInf.ClipText) > 0 Then

        sData = DataFimRealInf.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 185183

        DataFimRealInf.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFimReal_DownClick:

    Select Case gErr

        Case 185183

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185184)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFimReal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFimReal_UpClick

    DataFimRealInf.SetFocus

    If Len(Trim(DataFimRealInf.ClipText)) > 0 Then

        sData = DataFimRealInf.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 185185

        DataFimRealInf.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFimReal_UpClick:

    Select Case gErr

        Case 185185

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185186)

    End Select

    Exit Sub

End Sub

Private Sub DataFimRealInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFimRealInf, iAlterado)
    
End Sub

Private Sub DataFimRealInf_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFimRealInf_Validate

    If Len(Trim(DataFimRealInf.ClipText)) <> 0 Then

        lErro = Data_Critica(DataFimRealInf.Text)
        If lErro <> SUCESSO Then gError 185187
        
        PercCompRealInf.Text = "100"

    End If

    Exit Sub

Erro_DataFimRealInf_Validate:

    Cancel = True

    Select Case gErr

        Case 185187

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185188)

    End Select

    Exit Sub

End Sub

Private Sub DataFimRealInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercCompRealInf_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercCompRealInf_Validate

    'Verifica se PercCompRealInf está preenchida
    If Len(Trim(PercCompRealInf.Text)) <> 0 Then

       'Critica a PercCompRealInf
       lErro = Porcentagem_Critica(PercCompRealInf.Text)
       If lErro <> SUCESSO Then gError 185189

    End If

    Exit Sub

Erro_PercCompRealInf_Validate:

    Cancel = True

    Select Case gErr

        Case 185189

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185190)

    End Select

    Exit Sub

End Sub

Private Sub PercCompRealInf_GotFocus()
    Call MaskEdBox_TrataGotFocus(PercCompRealInf, iAlterado)
End Sub

Private Sub PercCompRealInf_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objEtapas As New ClassPRJEtapas
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Caption)) <> 0 Then

        objEtapas.sCodigo = Codigo.Caption

    End If

    Call Chama_Tela("PRJEtapasLista", colSelecao, objEtapas, objEventoCodigo, , "Código")

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185191)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeReduzido_Click()

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeReduzido_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(NomeReduzido.Text)) <> 0 Then

        objEtapa.sNomeReduzido = NomeReduzido.Text

    End If

    Call Chama_Tela("PRJEtapasLista", colSelecao, objEtapa, objEventoCodigo, , "Nome Reduzido")

    Exit Sub

Erro_LabelNomeReduzido_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185192)

    End Select

    Exit Sub

End Sub

Private Sub LabelReferencia_Click()

Dim lErro As Long
Dim objEtapas As New ClassPRJEtapas
Dim colSelecao As New Collection

On Error GoTo Erro_LabelReferencia_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Caption)) <> 0 Then

        objEtapas.sReferencia = Codigo.Caption

    End If

    Call Chama_Tela("PRJEtapasLista", colSelecao, objEtapas, objEventoCodigo, , "Referência")

    Exit Sub

Erro_LabelReferencia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185193)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objEtapa As ClassPRJEtapas

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objEtapa = obj1

    'Mostra os dados do CentrodeTrabalho na tela
    lErro = Traz_Etapa_Tela(objEtapa)
    If lErro <> SUCESSO Then gError 185194
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 185194

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185195)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)


End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(Opcao.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
    End If
    
    Call Atualiza_Indice

End Sub

'##################################################################
'Tem que colocar o código para o modo de edição aqui
Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReduzido, Source, X, Y)
End Sub

Private Sub LabelNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeRedPRJ_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeRedPRJ, Source, X, Y)
End Sub

Private Sub LabelNomeRedPRJ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeRedPRJ, Button, Shift, X, Y)
End Sub

Private Sub LabelReferencia_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelReferencia, Source, X, Y)
End Sub

Private Sub LabelReferencia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelReferencia, Button, Shift, X, Y)
End Sub
'##################################################################

'##################################################################
'Tratamento dos Campos customizados
Private Sub Data_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Data_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Data_GotFocus(Index)
End Sub

Private Sub Data_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Data_Validate(Index, Cancel)
End Sub

Private Sub Numero_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Numero_GotFocus(Index)
End Sub

Private Sub Numero_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Numero_Validate(Index, Cancel)
End Sub

Private Sub Valor_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Valor_GotFocus(Index As Integer)
    Call gobjTelaCamposCust.Valor_GotFocus(Index)
End Sub

Private Sub Valor_Validate(Index As Integer, Cancel As Boolean)
    Call gobjTelaCamposCust.Valor_Validate(Index, Cancel)
End Sub

Private Sub Texto_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownData_DownClick(Index As Integer)
    Call gobjTelaCamposCust.UpDownData_DownClick(Index)
End Sub

Private Sub UpDownData_UpClick(Index As Integer)
    Call gobjTelaCamposCust.UpDownData_UpClick(Index)
End Sub

Private Sub BotaoDadosCustNovo_Click()
    Call gobjTelaCamposCust.BotaoDadosCustNovo_Click
End Sub

Private Sub BotaoDadosCustDel_Click()
    Call gobjTelaCamposCust.BotaoDadosCustDel_Click
End Sub
'##################################################################

'##################################################################
'Tratamento do Escopo
Private Sub EscDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EscExclusoes_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EscExpectativa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EscFatores_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EscPremissas_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EscRestricoes_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'##################################################################

'##################################################################
'Tratamento Tab de Datas
Private Sub Intervalo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Intervalo_Validate(Cancel As Boolean)

Dim dtData As Date

    If Len(Trim(Intervalo)) <> 0 Then

        If StrParaDate(DataInicioInf.Text) <> DATA_NULA Then
                
            dtData = StrParaDate(DataInicioInf.Text)
            
            DataFimInf.PromptInclude = False
            DataFimInf.Text = Format(DateAdd("d", StrParaInt(Intervalo.Text) - 1, dtData), "dd/mm/yy")
            DataFimInf.PromptInclude = True
        
        Else
        
            If StrParaDate(DataFimInf.Text) <> DATA_NULA Then
    
                dtData = StrParaDate(DataFimInf.Text)
    
                DataInicioInf.PromptInclude = False
                DataInicioInf.Text = Format(DateAdd("d", -(StrParaInt(Intervalo.Text) - 1), dtData), "dd/mm/yy")
                DataInicioInf.PromptInclude = True
        
            End If
            
        End If
        
    End If
    
End Sub
'##################################################################

Private Function Inicializa_GridPred(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Etapa")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Data de Início")
    objGrid.colColuna.Add ("Data de Fim")

    'Controles que participam do Grid
    objGrid.colCampo.Add (PredCodigo.Name)
    objGrid.colCampo.Add (PredDescricao.Name)
    objGrid.colCampo.Add (PredDataIni.Name)
    objGrid.colCampo.Add (PredDataFim.Name)

    'Colunas do Grid
    iGrid_PredCodigo_Col = 1
    iGrid_PredDescricao_Col = 2
    iGrid_PredDataIni_Col = 3
    iGrid_PredDataFim_Col = 4

    objGrid.objGrid = GridPred

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridPred.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridPred = SUCESSO

End Function

Private Sub GridPred_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPred, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPred, iAlterado)
    End If

End Sub

Private Sub GridPred_GotFocus()
    Call Grid_Recebe_Foco(objGridPred)
End Sub

Private Sub GridPred_EnterCell()
    Call Grid_Entrada_Celula(objGridPred, iAlterado)
End Sub

Private Sub GridPred_LeaveCell()
    Call Saida_Celula(objGridPred)
End Sub

Private Sub GridPred_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPred, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPred, iAlterado)
    End If

End Sub

Private Sub GridPred_RowColChange()
    Call Grid_RowColChange(objGridPred)
End Sub

Private Sub GridPred_Scroll()
    Call Grid_Scroll(objGridPred)
End Sub

Private Sub GridPred_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridPred)
End Sub

Private Sub GridPred_LostFocus()
    Call Grid_Libera_Foco(objGridPred)
End Sub

Private Function Inicializa_GridPA(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (PAProduto.Name)
    objGrid.colCampo.Add (PADescricao.Name)
    objGrid.colCampo.Add (PAVersao.Name)
    objGrid.colCampo.Add (PAUM.Name)
    objGrid.colCampo.Add (PAQuantidade.Name)
    objGrid.colCampo.Add (PAObservacao.Name)

    'Colunas do Grid
    iGrid_PAProduto_Col = 1
    iGrid_PADescricao_Col = 2
    iGrid_PAVersao_Col = 3
    iGrid_PAUM_Col = 4
    iGrid_PAQuantidade_Col = 5
    iGrid_PAObservacao_Col = 6

    objGrid.objGrid = GridPA

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridPA.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridPA = SUCESSO

End Function

Private Sub GridPA_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPA, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPA, iAlterado)
    End If

End Sub

Private Sub GridPA_GotFocus()
    Call Grid_Recebe_Foco(objGridPA)
End Sub

Private Sub GridPA_EnterCell()
    Call Grid_Entrada_Celula(objGridPA, iAlterado)
End Sub

Private Sub GridPA_LeaveCell()
    Call Saida_Celula(objGridPA)
End Sub

Private Sub GridPA_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPA, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPA, iAlterado)
    End If

End Sub

Private Sub GridPA_RowColChange()
    Call Grid_RowColChange(objGridPA)
End Sub

Private Sub GridPA_Scroll()
    Call Grid_Scroll(objGridPA)
End Sub

Private Sub GridPA_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

On Error GoTo Erro_GridPA_KeyDown

    'Guarda iLinhasExistentes
    iLinhasExistentesAnterior = objGridPA.iLinhasExistentes
    
    'Verifica se a Tecla apertada foi Del
    If KeyCode = vbKeyDelete Then
        'Guarda o índice da Linha a ser Excluída
        iLinhaAnterior = GridPA.Row
    End If

    Call Grid_Trata_Tecla1(KeyCode, objGridPA)

    'Verifica se a Linha foi realmente excluída
    If objGridPA.iLinhasExistentes < iLinhasExistentesAnterior Then

        lErro = Recalcula_Necessidades
        If lErro <> SUCESSO Then gError 185449
        
    End If
    
    Exit Sub
    
Erro_GridPA_KeyDown:

    Select Case gErr
    
        Case 185449
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185450)
    
    End Select

    Exit Sub
    
End Sub

Private Sub GridPA_LostFocus()
    Call Grid_Libera_Foco(objGridPA)
End Sub

Private Function Inicializa_GridMP(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colColuna.Add ("Data")
    End If
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Custo")
    objGrid.colColuna.Add ("Custo Total")
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colColuna.Add ("Observação")
    'End If
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (MPProduto(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MPDescricao(iIndiceInfCalcAtual).Name)
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colCampo.Add (MPData(iIndiceInfCalcAtual).Name)
    End If
    objGrid.colCampo.Add (MPUM(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MPQuantidade(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MPCusto(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MPCustoT(iIndiceInfCalcAtual).Name)
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colCampo.Add (MPOBS(iIndiceInfCalcAtual).Name)
    'End If
    
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colIndex.Add iIndiceInfCalcAtual
    End If
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colIndex.Add iIndiceInfCalcAtual
    'End If
    
    'Colunas do Grid
    iIndice = 0
    iGrid_MPProduto_Col(iIndiceInfCalcAtual) = 1 + iIndice
    iGrid_MPDescricao_Col(iIndiceInfCalcAtual) = 2 + iIndice
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        iGrid_MPData_Col(iIndiceInfCalcAtual) = 3 + iIndice
        iIndice = iIndice + 1
    End If
    iGrid_MPUM_Col(iIndiceInfCalcAtual) = 3 + iIndice
    iGrid_MPQuantidade_Col(iIndiceInfCalcAtual) = 4 + iIndice
    iGrid_MPCusto_Col(iIndiceInfCalcAtual) = 5 + iIndice
    iGrid_MPCustoT_Col(iIndiceInfCalcAtual) = 6 + iIndice
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        iGrid_MPOBS_Col(iIndiceInfCalcAtual) = 7 + iIndice
        iIndice = iIndice + 1
    'End If
    
    objGrid.objGrid = GridMP(iIndiceInfCalcAtual)

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 9
    
    If iIndiceInfCalcAtual = INDICE_CALC_PREV Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
        objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    End If

    'Largura da primeira coluna
    GridMP(iIndiceInfCalcAtual).ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMP = SUCESSO

End Function

Private Sub GridMP_Click(iIndex As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMP(iIndex), iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMP(iIndex), iAlterado)
    End If

End Sub

Private Sub GridMP_GotFocus(iIndex As Integer)
    Call Grid_Recebe_Foco(objGridMP(iIndex))
End Sub

Private Sub GridMP_EnterCell(iIndex As Integer)
    Call Grid_Entrada_Celula(objGridMP(iIndex), iAlterado)
End Sub

Private Sub GridMP_LeaveCell(iIndex As Integer)
    Call Saida_Celula(objGridMP(iIndex))
End Sub

Private Sub GridMP_KeyPress(iIndex As Integer, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMP(iIndex), iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMP(iIndex), iAlterado)
    End If

End Sub

Private Sub GridMP_RowColChange(iIndex As Integer)
    Call Grid_RowColChange(objGridMP(iIndex))
End Sub

Private Sub GridMP_Scroll(iIndex As Integer)
    Call Grid_Scroll(objGridMP(iIndex))
End Sub

Private Sub GridMP_KeyDown(iIndex As Integer, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMP(iIndex))

    Call Soma_Coluna_Grid(objGridMP(iIndex), iGrid_MPCustoT_Col(iIndex), MPCustoTotal(iIndex))
End Sub

Private Sub GridMP_LostFocus(iIndex As Integer)
    Call Grid_Libera_Foco(objGridMP(iIndex))
End Sub

Private Function Inicializa_GridMaq(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Máquina")
    objGrid.colColuna.Add ("Descrição")
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colColuna.Add ("Data")
    End If
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Horas")
    objGrid.colColuna.Add ("Custo UN/h")
    objGrid.colColuna.Add ("Custo Total")
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colColuna.Add ("Observação")
    'End If
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (MaqCodigo(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MaqDescricao(iIndiceInfCalcAtual).Name)
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colCampo.Add (MaqData(iIndiceInfCalcAtual).Name)
    End If
    objGrid.colCampo.Add (MaqQuantidade(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MaqHoras(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MaqCusto(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MaqCustoT(iIndiceInfCalcAtual).Name)
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colCampo.Add (MaqOBS(iIndiceInfCalcAtual).Name)
    'End If
    
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colIndex.Add iIndiceInfCalcAtual
    End If
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colIndex.Add iIndiceInfCalcAtual
    'End If
    
    'Colunas do Grid
    iIndice = 0
    iGrid_MaqCodigo_Col(iIndiceInfCalcAtual) = 1 + iIndice
    iGrid_MaqDescricao_Col(iIndiceInfCalcAtual) = 2 + iIndice
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        iGrid_MaqData_Col(iIndiceInfCalcAtual) = 3 + iIndice
        iIndice = iIndice + 1
    End If
    iGrid_MaqQuantidade_Col(iIndiceInfCalcAtual) = 3 + iIndice
    iGrid_MaqHoras_Col(iIndiceInfCalcAtual) = 4 + iIndice
    iGrid_MaqCusto_Col(iIndiceInfCalcAtual) = 5 + iIndice
    iGrid_MaqCustoT_Col(iIndiceInfCalcAtual) = 6 + iIndice
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        iGrid_MaqOBS_Col(iIndiceInfCalcAtual) = 7 + iIndice
        iIndice = iIndice + 1
    'End If
    
    objGrid.objGrid = GridMaq(iIndiceInfCalcAtual)

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 9
    
    If iIndiceInfCalcAtual = INDICE_CALC_PREV Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
        objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    End If

    'Largura da primeira coluna
    GridMaq(iIndiceInfCalcAtual).ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaq = SUCESSO

End Function

Private Sub GridMaq_Click(iIndex As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaq(iIndex), iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaq(iIndex), iAlterado)
    End If

End Sub

Private Sub GridMaq_GotFocus(iIndex As Integer)
    Call Grid_Recebe_Foco(objGridMaq(iIndex))
End Sub

Private Sub GridMaq_EnterCell(iIndex As Integer)
    Call Grid_Entrada_Celula(objGridMaq(iIndex), iAlterado)
End Sub

Private Sub GridMaq_LeaveCell(iIndex As Integer)
    Call Saida_Celula(objGridMaq(iIndex))
End Sub

Private Sub GridMaq_KeyPress(iIndex As Integer, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMaq(iIndex), iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaq(iIndex), iAlterado)
    End If

End Sub

Private Sub GridMaq_RowColChange(iIndex As Integer)
    Call Grid_RowColChange(objGridMaq(iIndex))
End Sub

Private Sub GridMaq_Scroll(iIndex As Integer)
    Call Grid_Scroll(objGridMaq(iIndex))
End Sub

Private Sub GridMaq_KeyDown(iIndex As Integer, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMaq(iIndex))
    
    Call Soma_Coluna_Grid(objGridMaq(iIndex), iGrid_MaqCustoT_Col(iIndex), MaqCustoTotal(iIndex))

End Sub

Private Sub GridMaq_LostFocus(iIndex As Integer)
    Call Grid_Libera_Foco(objGridMaq(iIndex))
End Sub

Private Function Inicializa_GridMO(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Mão de Obra")
    objGrid.colColuna.Add ("Descrição")
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colColuna.Add ("Data")
    End If
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Horas")
    objGrid.colColuna.Add ("Custo UN/h")
    objGrid.colColuna.Add ("Custo Total")
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colColuna.Add ("Observação")
    'End If
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (MOCodigo(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MODescricao(iIndiceInfCalcAtual).Name)
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colCampo.Add (MOData(iIndiceInfCalcAtual).Name)
    End If
    objGrid.colCampo.Add (MOQuantidade(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MOHoras(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MOCusto(iIndiceInfCalcAtual).Name)
    objGrid.colCampo.Add (MOCustoT(iIndiceInfCalcAtual).Name)
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colCampo.Add (MOOBS(iIndiceInfCalcAtual).Name)
    'End If
    
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colIndex.Add iIndiceInfCalcAtual
    End If
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    objGrid.colIndex.Add iIndiceInfCalcAtual
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.colIndex.Add iIndiceInfCalcAtual
    'End If
    
    'Colunas do Grid
    iIndice = 0
    iGrid_MOCodigo_Col(iIndiceInfCalcAtual) = 1 + iIndice
    iGrid_MODescricao_Col(iIndiceInfCalcAtual) = 2 + iIndice
    If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        iGrid_MOData_Col(iIndiceInfCalcAtual) = 3 + iIndice
        iIndice = iIndice + 1
    End If
    iGrid_MOQuantidade_Col(iIndiceInfCalcAtual) = 3 + iIndice
    iGrid_MOHoras_Col(iIndiceInfCalcAtual) = 4 + iIndice
    iGrid_MOCusto_Col(iIndiceInfCalcAtual) = 5 + iIndice
    iGrid_MOCustoT_Col(iIndiceInfCalcAtual) = 6 + iIndice
    'If iIndiceInfCalcAtual = INDICE_INF_REAL Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        iGrid_MOOBS_Col(iIndiceInfCalcAtual) = 7 + iIndice
        iIndice = iIndice + 1
    'End If
    
    objGrid.objGrid = GridMO(iIndiceInfCalcAtual)

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 9
    
    If iIndiceInfCalcAtual = INDICE_CALC_PREV Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
        objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    End If

    'Largura da primeira coluna
    GridMO(iIndiceInfCalcAtual).ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMO = SUCESSO

End Function

Private Sub GridMO_Click(iIndex As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMO(iIndex), iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO(iIndex), iAlterado)
    End If

End Sub

Private Sub GridMO_GotFocus(iIndex As Integer)
    Call Grid_Recebe_Foco(objGridMO(iIndex))
End Sub

Private Sub GridMO_EnterCell(iIndex As Integer)
    Call Grid_Entrada_Celula(objGridMO(iIndex), iAlterado)
End Sub

Private Sub GridMO_LeaveCell(iIndex As Integer)
    Call Saida_Celula(objGridMO(iIndex))
End Sub

Private Sub GridMO_KeyPress(iIndex As Integer, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMO(iIndex), iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO(iIndex), iAlterado)
    End If

End Sub

Private Sub GridMO_RowColChange(iIndex As Integer)
    Call Grid_RowColChange(objGridMO(iIndex))
End Sub

Private Sub GridMO_Scroll(iIndex As Integer)
    Call Grid_Scroll(objGridMO(iIndex))
End Sub

Private Sub GridMO_KeyDown(iIndex As Integer, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMO(iIndex))

    Call Soma_Coluna_Grid(objGridMO(iIndex), iGrid_MOCustoT_Col(iIndex), MOCustoTotal(iIndex))
End Sub

Private Sub GridMO_LostFocus(iIndex As Integer)
    Call Grid_Libera_Foco(objGridMO(iIndex))
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    If iIndiceInfCalcAtual = 0 Then iIndiceInfCalcAtual = 1

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridPred.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_PredCodigo_Col
                
                    lErro = Saida_Celula_PredCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 185196
        
            End Select
                
        ElseIf objGridInt.objGrid.Name = GridPA.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_PAProduto_Col
                
                    lErro = Saida_Celula_PAProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 185197

                Case iGrid_PADescricao_Col
                
                    lErro = Saida_Celula_PADescricao(objGridInt)
                    If lErro <> SUCESSO Then gError 185198

                Case iGrid_PAVersao_Col
                
                    lErro = Saida_Celula_PAVersao(objGridInt)
                    If lErro <> SUCESSO Then gError 185199

                Case iGrid_PAQuantidade_Col
                
                    lErro = Saida_Celula_PAQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 185200

                Case iGrid_PAObservacao_Col
                
                    lErro = Saida_Celula_PAObservacao(objGridInt)
                    If lErro <> SUCESSO Then gError 185201
                    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridMP(iIndiceInfCalcAtual).Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
            
                Case iGrid_MPProduto_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MPProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 185202

                Case iGrid_MPDescricao_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MPDescricao(objGridInt)
                    If lErro <> SUCESSO Then gError 185203

                Case iGrid_MPQuantidade_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MPQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 185204

                Case iGrid_MPCusto_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MPCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 185205

                Case iGrid_MPData_Col(iIndiceInfCalcAtual)
                
                    If iGrid_MPData_Col(iIndiceInfCalcAtual) <> 0 Then
                        lErro = Saida_Celula_Data(objGridInt, MPData(iIndiceInfCalcAtual))
                        If lErro <> SUCESSO Then gError 185205
                    End If

                Case iGrid_MPOBS_Col(iIndiceInfCalcAtual)
                
                    If iGrid_MPOBS_Col(iIndiceInfCalcAtual) <> 0 Then
                        lErro = Saida_Celula_Padrao(objGridInt, MPOBS(iIndiceInfCalcAtual))
                        If lErro <> SUCESSO Then gError 185205
                    End If

            End Select
                    
                
        ElseIf objGridInt.objGrid.Name = GridMO(iIndiceInfCalcAtual).Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_MOCodigo_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MOCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 185207

                Case iGrid_MODescricao_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MODescricao(objGridInt)
                    If lErro <> SUCESSO Then gError 185208

                Case iGrid_MOQuantidade_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MOQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 185209

                Case iGrid_MOHoras_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MOHoras(objGridInt)
                    If lErro <> SUCESSO Then gError 185210

                Case iGrid_MOCusto_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MOCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 185211
                    
                Case iGrid_MOData_Col(iIndiceInfCalcAtual)
                
                    If iGrid_MOData_Col(iIndiceInfCalcAtual) <> 0 Then
                        lErro = Saida_Celula_Data(objGridInt, MOData(iIndiceInfCalcAtual))
                        If lErro <> SUCESSO Then gError 185205
                    End If

                Case iGrid_MOOBS_Col(iIndiceInfCalcAtual)
                
                    If iGrid_MOOBS_Col(iIndiceInfCalcAtual) <> 0 Then
                        lErro = Saida_Celula_Padrao(objGridInt, MOOBS(iIndiceInfCalcAtual))
                        If lErro <> SUCESSO Then gError 185205
                    End If
        
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridMaq(iIndiceInfCalcAtual).Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_MaqCodigo_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MaqCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 185212

                Case iGrid_MaqDescricao_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MaqDescricao(objGridInt)
                    If lErro <> SUCESSO Then gError 185213

                Case iGrid_MaqQuantidade_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MaqQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 185214

                Case iGrid_MaqHoras_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MaqHoras(objGridInt)
                    If lErro <> SUCESSO Then gError 185215

                Case iGrid_MaqCusto_Col(iIndiceInfCalcAtual)
                
                    lErro = Saida_Celula_MaqCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 185216
                    
                Case iGrid_MaqData_Col(iIndiceInfCalcAtual)
                
                    If iGrid_MaqData_Col(iIndiceInfCalcAtual) <> 0 Then
                        lErro = Saida_Celula_Data(objGridInt, MaqData(iIndiceInfCalcAtual))
                        If lErro <> SUCESSO Then gError 185205
                    End If

                Case iGrid_MaqOBS_Col(iIndiceInfCalcAtual)
                
                    If iGrid_MaqOBS_Col(iIndiceInfCalcAtual) <> 0 Then
                        lErro = Saida_Celula_Padrao(objGridInt, MaqOBS(iIndiceInfCalcAtual))
                        If lErro <> SUCESSO Then gError 185205
                    End If
                    
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 185217

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 185196 To 185216
            'erros tratatos nas rotinas chamadas
        
        Case 185217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185218)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatadoPA As String
Dim iProdutoPreenchidoPA As Integer
Dim sProdutoFormatadoMP As String
Dim iProdutoPreenchidoMP As Integer
Dim iMaquinaPreenchida As Integer
Dim iTipoMaoDeObraPreenchida As Integer
Dim objKit As ClassKit
Dim colKits As Collection
Dim iIndice As Integer
Dim sVersaoAnt As String
Dim iIndiceAnt As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    lErro = CF("Produto_Formata", GridPA.TextMatrix(GridPA.Row, iGrid_PAProduto_Col), sProdutoFormatadoPA, iProdutoPreenchidoPA)
    If lErro <> SUCESSO Then gError 185219

    iIndiceAnt = iIndiceInfCalcAtual

    If iIndiceInfCalcAtual <> 0 Then
    
        lErro = CF("Produto_Formata", GridMP(iIndiceInfCalcAtual).TextMatrix(GridMP(iIndiceInfCalcAtual).Row, iGrid_MPProduto_Col(iIndiceInfCalcAtual)), sProdutoFormatadoMP, iProdutoPreenchidoMP)
        If lErro <> SUCESSO Then gError 185220
        
        If Len(Trim(GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCodigo_Col(iIndiceInfCalcAtual)))) > 0 Then
            iTipoMaoDeObraPreenchida = MARCADO
        Else
            iTipoMaoDeObraPreenchida = DESMARCADO
        End If
        
        If Len(Trim(GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCodigo_Col(iIndiceInfCalcAtual)))) > 0 Then
            iMaquinaPreenchida = MARCADO
        Else
            iMaquinaPreenchida = DESMARCADO
        End If
    
    Else
        iIndiceInfCalcAtual = 1
    End If
        
    Select Case objControl.Name
    
        Case PAProduto.Name
        
            If iProdutoPreenchidoPA = PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case PADescricao.Name, PAQuantidade.Name, PAObservacao.Name
            
            If iProdutoPreenchidoPA <> PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case PAUM.Name
            objControl.Enabled = False
            
        Case PAVersao.Name
                        
            If iProdutoPreenchidoPA <> PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            
                'Inserido por Wagner
                sVersaoAnt = PAVersao.Text
                
                PAVersao.Clear
    
                If iProdutoPreenchidoPA <> PRODUTO_VAZIO Then
                
                    Set objKit = New ClassKit
                    Set colKits = New Collection
                    
                    'Armazena o Produto Raiz do kit
                    objKit.sProdutoRaiz = sProdutoFormatadoPA
                    
                    'Le as Versoes Ativas e a Padrao
                    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
                    If lErro <> SUCESSO And lErro <> 106333 Then gError 185823
                    
                    'Carrega a Combo com os Dados da Colecao
                    For Each objKit In colKits
                    
                        PAVersao.AddItem (objKit.sVersao)
                                                    
                    Next
                    
                    'Tento selecionar na Combo a Unidade anterior
                    If PAVersao.ListCount <> 0 Then
        
                        For iIndice = 0 To PAVersao.ListCount - 1
        
                            If PAVersao.List(iIndice) = sVersaoAnt Then
                                PAVersao.ListIndex = iIndice
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
            
        Case MPProduto(iIndiceInfCalcAtual).Name
        
            If iProdutoPreenchidoMP = PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case MPDescricao(iIndiceInfCalcAtual).Name, MPQuantidade(iIndiceInfCalcAtual).Name, MPCusto(iIndiceInfCalcAtual).Name, "MPOBS", "MPData"

            If iProdutoPreenchidoMP <> PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case MPUM(iIndiceInfCalcAtual).Name
        
            objControl.Enabled = False
            
        Case MOCodigo(iIndiceInfCalcAtual).Name
        
            If iTipoMaoDeObraPreenchida = MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case MODescricao(iIndiceInfCalcAtual).Name, MOQuantidade(iIndiceInfCalcAtual).Name, MOCusto(iIndiceInfCalcAtual).Name, MOHoras(iIndiceInfCalcAtual).Name, "MOOBS", "MOData"

            If iTipoMaoDeObraPreenchida <> MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case MaqCodigo(iIndiceInfCalcAtual).Name
        
            If iMaquinaPreenchida = MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case MaqDescricao(iIndiceInfCalcAtual).Name, MaqQuantidade(iIndiceInfCalcAtual).Name, MaqCusto(iIndiceInfCalcAtual).Name, MaqHoras(iIndiceInfCalcAtual).Name, "MaqOBS", "MaqData"

            If iMaquinaPreenchida <> MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case PredCodigo.Name
            objControl.Enabled = True
            
        Case Else
            objControl.Enabled = False
            
    End Select
    
    iIndiceInfCalcAtual = iIndiceAnt
    
    If iIndiceInfCalcAtual = INDICE_CALC_PREV Or iIndiceInfCalcAtual = INDICE_CALC_REAL Then
        objControl.Enabled = False
    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 185219 To 185221, 185823

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185222)

    End Select

    Exit Sub

End Sub

Sub Projeto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim lNumIntDocPRJ As Long
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Projeto_Validate

    'Se alterou o projeto
    If sProjetoAnt <> Projeto.Text Then

        If Len(Trim(Projeto.ClipText)) > 0 Then
            
            lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError 189070
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185223
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185224
            
            lNumIntDocPRJ = objProjeto.lNumIntDoc
            
            NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
            
        End If
        
        sProjetoAnt = Projeto.Text
        sNomeProjetoAnt = NomeReduzidoPRJ.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 185225
        
    End If
   
    Exit Sub

Erro_Projeto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 185223, 185225, 189070
        
        Case 185224
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185226)

    End Select

    Exit Sub

End Sub

Sub NomeReduzidoPrj_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim lNumIntDocPRJ As Long

On Error GoTo Erro_NomeReduzidoPrj_Validate

    'Se alterou o projeto
    If sNomeProjetoAnt <> NomeReduzidoPRJ.Text Then

        If Len(Trim(NomeReduzidoPRJ.Text)) > 0 Then
            
            objProjeto.sNomeReduzido = NomeReduzidoPRJ.Text
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le_NomeReduzido", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185227
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 185228
            
            lNumIntDocPRJ = objProjeto.lNumIntDoc
            
            lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
            If lErro <> SUCESSO Then gError 189109
            
        End If
        
        sProjetoAnt = Projeto.Text
        sNomeProjetoAnt = NomeReduzidoPRJ.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 185229
        
    End If
    
    Exit Sub

Erro_NomeReduzidoPrj_Validate:

    Cancel = True

    Select Case gErr
    
        Case 185227, 185229, 189109
        
        Case 185228
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO3", gErr, objProjeto.sNomeReduzido, objProjeto.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185230)

    End Select

    Exit Sub

End Sub

Sub LabelProjeto_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjeto_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Projeto.ClipText)) <> 0 Then

        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189071

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoPRJ, , "Código")

    Exit Sub

Erro_LabelProjeto_Click:

    Select Case gErr
    
        Case 189071

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185231)

    End Select

    Exit Sub
    
End Sub

Sub LabelNomeRedPRJ_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeRedPRJ_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(NomeReduzidoPRJ.Text)) <> 0 Then

        objProjeto.sNomeReduzido = NomeReduzidoPRJ.Text

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoPRJ, , "Nome Reduzido")

    Exit Sub

Erro_LabelNomeRedPRJ_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185232)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoPRJ_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoPRJ_evSelecao

    Set objProjeto = obj1

    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189110
    
    NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
    
    Call Projeto_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoPRJ_evSelecao:

    Select Case gErr
    
        Case 189110

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185233)

    End Select

    Exit Sub

End Sub

Private Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomeReduzidoPrj_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Referencia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PredCodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PredCodigo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PredCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPred)
End Sub

Private Sub PredCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPred)
End Sub

Private Sub PredCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPred.objControle = PredCodigo
    lErro = Grid_Campo_Libera_Foco(objGridPred)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Trata_Projeto(ByVal lNumIntDocPRJ As Long) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos

On Error GoTo Erro_Trata_Projeto

    Codigo.Caption = ""
    
    Call Grid_Limpa(objGridPred)
    
    If lNumIntDocPRJ <> 0 Then

        objProjeto.lNumIntDoc = lNumIntDocPRJ
    
        lErro = CF("CarregaCombo_Etapas", objProjeto, PredCodigo)
        If lErro <> SUCESSO Then gError 185234
        
    Else
    
        PredCodigo.Clear
        
    End If

    Trata_Projeto = SUCESSO

    Exit Function

Erro_Trata_Projeto:

    Trata_Projeto = gErr

    Select Case gErr
    
        Case 185234

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185235)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PredCodigo(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim objEtapa As New ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Saida_Celula_PredCodigo

    Set objGridInt.objControle = PredCodigo
    
    'Se o campo foi preenchido
    If Len(Trim(PredCodigo.Text)) > 0 Then
    
        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189073
    
        objProjeto.sCodigo = sProjeto
        objProjeto.iFilialEmpresa = giFilialEmpresa

        'Le o projeto
        lErro = CF("Projetos_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185236
        
        objEtapa.sCodigo = SCodigo_Extrai(PredCodigo.Text)
        objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
        
        If objEtapa.sCodigo = Codigo.Caption Then gError 185485

        'Le a etapa
        lErro = CF("PRJEtapas_Le", objEtapa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 185237
        
        If objEtapa.dtDataFim <> DATA_NULA Then
            GridPred.TextMatrix(GridPred.Row, iGrid_PredDataFim_Col) = Format(objEtapa.dtDataFim, "dd/mm/yyyy")
        Else
            GridPred.TextMatrix(GridPred.Row, iGrid_PredDataFim_Col) = ""
        End If
        
        If objEtapa.dtDataInicio <> DATA_NULA Then
            GridPred.TextMatrix(GridPred.Row, iGrid_PredDataIni_Col) = Format(objEtapa.dtDataInicio, "dd/mm/yyyy")
        Else
            GridPred.TextMatrix(GridPred.Row, iGrid_PredDataIni_Col) = ""
        End If
        
        GridPred.TextMatrix(GridPred.Row, iGrid_PredDescricao_Col) = objEtapa.sDescricao
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridPred.Row - GridPred.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
    
        GridPred.TextMatrix(GridPred.Row, iGrid_PredDataFim_Col) = ""
        GridPred.TextMatrix(GridPred.Row, iGrid_PredDataIni_Col) = ""
        GridPred.TextMatrix(GridPred.Row, iGrid_PredDescricao_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185238

    Saida_Celula_PredCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_PredCodigo:

    Saida_Celula_PredCodigo = gErr

    Select Case gErr

        Case 185236 To 185238, 189073
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 185485
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_PRED_SI_MESMA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185239)

    End Select

    Exit Function

End Function

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridPA.Row = 0 Then gError 185240

    'Verifica se o Produto está preenchido
    If Len(Trim(GridPA.TextMatrix(GridPA.Row, iGrid_PAProduto_Col))) > 0 Then
    
        lErro = CF("Produto_Formata", GridPA.TextMatrix(GridPA.Row, iGrid_PAProduto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 185241
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
    End If

    objProduto.sCodigo = sProduto

    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoPA)
        
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 185240
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 185241
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185242)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoPA_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objItemOP As New ClassItemOP

On Error GoTo Erro_objEventoPA_evSelecao

    Set objProduto = obj1

    If GridPA.Row <> 0 Then

        lErro = CF("Produto_Formata", GridPA.TextMatrix(GridPA.Row, iGrid_PAProduto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 185243

        'Se o produto não estiver preenchido
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            'Lê o produto no BD para obter UM de estoque
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 185244

            If lErro = 28030 Then gError 185245

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            'mascara produto escolhido
            lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 185246

            PAProduto.PromptInclude = False
            PAProduto.Text = sProdutoMascarado
            PAProduto.PromptInclude = True
            
            If Not (Me.ActiveControl Is PAProduto) Then

                'preenche produto
                GridPA.TextMatrix(GridPA.Row, iGrid_PAProduto_Col) = sProdutoMascarado

                'Preenche a linha do grid
                lErro = ProdutoLinha_PreenchePA(objProduto)
                If lErro <> SUCESSO Then gError 185247

            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoPA_evSelecao:

    Select Case gErr

        Case 185243, 185244, 185247
           
        Case 185245
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 185246
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185248)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_PreenchePA(objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoLinha_PreenchePA

    PAVersao.Clear

    Call Carrega_ComboVersoes(objProduto.sCodigo)

    'se o produto nao controla estoque ==> erro
    If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 185249

    If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 185250

    'Unidade de Medida
    GridPA.TextMatrix(GridPA.Row, iGrid_PAUM_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridPA.TextMatrix(GridPA.Row, iGrid_PADescricao_Col) = objProduto.sDescricao

    'Versão
    lErro = Carrega_ComboVersoes(objProduto.sCodigo)
    If lErro <> SUCESSO Then gError 185251

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridPA.Row - GridPA.FixedRows) = objGridPA.iLinhasExistentes Then
        objGridPA.iLinhasExistentes = objGridPA.iLinhasExistentes + 1
    End If

    ProdutoLinha_PreenchePA = SUCESSO

    Exit Function

Erro_ProdutoLinha_PreenchePA:

    ProdutoLinha_PreenchePA = gErr

    Select Case gErr
                
        Case 185249
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)

        Case 185250
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, PAProduto.Text)

        Case 185251

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185252)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_PreencheMP(objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoLinha_PreencheMP

    GridMP(iIndiceInfCalcAtual).TextMatrix(GridMP(iIndiceInfCalcAtual).Row, iGrid_MPDescricao_Col(iIndiceInfCalcAtual)) = objProduto.sDescricao
    GridMP(iIndiceInfCalcAtual).TextMatrix(GridMP(iIndiceInfCalcAtual).Row, iGrid_MPUM_Col(iIndiceInfCalcAtual)) = objProduto.sSiglaUMEstoque
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridMP(iIndiceInfCalcAtual).Row - GridMP(iIndiceInfCalcAtual).FixedRows = objGridMP(iIndiceInfCalcAtual).iLinhasExistentes Then
        objGridMP(iIndiceInfCalcAtual).iLinhasExistentes = objGridMP(iIndiceInfCalcAtual).iLinhasExistentes + 1
    End If

    ProdutoLinha_PreencheMP = SUCESSO

    Exit Function

Erro_ProdutoLinha_PreencheMP:

    ProdutoLinha_PreencheMP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 185253)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboVersoes(ByVal sProdutoRaiz As String) As Long
    
Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
Dim iPadrao As Integer
Dim iIndice As Integer
Dim sVersaoAnt As String
Dim bAchou As Boolean
    
On Error GoTo Erro_Carrega_ComboVersoes
    
    'Limpa a Combo
    PAVersao.Clear
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProdutoRaiz
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 185254
    
    iPadrao = -1
    
    'Carrega a Combo com os Dados da Colecao
    For Each objKit In colKits
    
        PAVersao.AddItem (objKit.sVersao)
        
        'Se for a padrao -> Armazena
        If objKit.iSituacao = KIT_SITUACAO_PADRAO Then iPadrao = iIndice
        
        iIndice = iIndice + 1
        
    Next
    
    If Len(GridPA.TextMatrix(GridPA.Row, iGrid_PAVersao_Col)) > 0 Then
    
        PAVersao.Text = GridPA.TextMatrix(GridPA.Row, iGrid_PAVersao_Col)
    
    ElseIf iPadrao <> -1 Then
    
        'Seleciona a Padrao na Combo
        PAVersao.ListIndex = iPadrao
        
        GridPA.TextMatrix(GridPA.Row, iGrid_PAVersao_Col) = PAVersao.Text

    End If
    
    Carrega_ComboVersoes = SUCESSO

    Exit Function
    
Erro_Carrega_ComboVersoes:

    Carrega_ComboVersoes = gErr

    Select Case gErr
    
        Case 185254
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185255)
    
    End Select
    
End Function

Private Function Saida_Celula_PAVersao(objGridInt As AdmGrid) As Long
'faz a critica da celula de Versao do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Saida_Celula_PAVersao

    Set objGridInt.objControle = PAVersao
    
    'Se a quantidade estiver preenchida
    If StrParaDbl(GridPA.TextMatrix(GridPA.Row, iGrid_PAQuantidade_Col)) > 0 Then
    
        'Se alteroo a versão
        If GridPA.TextMatrix(GridPA.Row, iGrid_PAVersao_Col) <> PAVersao.Text Then
            lErro = Recalcula_Necessidades
            If lErro <> SUCESSO Then gError 185448
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185256

    Saida_Celula_PAVersao = SUCESSO

    Exit Function

Erro_Saida_Celula_PAVersao:

    Saida_Celula_PAVersao = gErr

    Select Case gErr

        Case 185256, 185448
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185257)

    End Select

End Function

Private Function Saida_Celula_PAQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantAnt As Double

On Error GoTo Erro_Saida_Celula_PAQuantidade

    Set objGridInt.objControle = PAQuantidade
    
    'se a quantidade foi preenchida
    If Len(PAQuantidade.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(PAQuantidade.Text)
        If lErro <> SUCESSO Then gError 185258
    
        PAQuantidade.Text = Formata_Estoque(PAQuantidade.Text)
    
    End If
    
    dQuantAnt = StrParaDbl(GridPA.TextMatrix(GridPA.Row, iGrid_PAQuantidade_Col))
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185259
    
    If Abs(dQuantAnt - StrParaDbl(PAQuantidade.Text)) > QTDE_ESTOQUE_DELTA Then
        lErro = Recalcula_Necessidades
        If lErro <> SUCESSO Then gError 185445
    End If

    Saida_Celula_PAQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_PAQuantidade:

    Saida_Celula_PAQuantidade = gErr

    Select Case gErr
            
        Case 185258
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            PAQuantidade.SetFocus
            
        Case 185259, 185445
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185260)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PAObservacao(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PAObservacao

    Set objGridInt.objControle = PAObservacao
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185261

    Saida_Celula_PAObservacao = SUCESSO

    Exit Function

Erro_Saida_Celula_PAObservacao:

    Saida_Celula_PAObservacao = gErr

    Select Case gErr

        Case 185261
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185262)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PADescricao(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PADescricao

    Set objGridInt.objControle = PADescricao
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185263

    Saida_Celula_PADescricao = SUCESSO

    Exit Function

Erro_Saida_Celula_PADescricao:

    Saida_Celula_PADescricao = gErr

    Select Case gErr

        Case 185263
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185264)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PAProduto(objGridInt As AdmGrid) As Long
'faz a critica da celula de proddduto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim sProdutoMascarado As String

On Error GoTo Erro_Saida_Celula_PAProduto

    Set objGridInt.objControle = PAProduto

    lErro = CF("Produto_Formata", PAProduto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 185265
    
    'se o produto foi preenchido
    If Len(Trim(PAProduto.ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica2", PAProduto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 185266
        
        'mascara produto escolhido
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 185267

        PAProduto.PromptInclude = False
        PAProduto.Text = sProdutoMascarado
        PAProduto.PromptInclude = True
        
        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
            'se é um produto gerencial e não é pai de grade ==> erro
            If lErro = 25043 And Len(Trim(objProduto.sGrade)) = 0 Then gError 185268
            
            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
            If lErro <> SUCESSO And lErro <> 25043 Then gError 185269
        
             'Preenche a linha do grid
            lErro = ProdutoLinha_PreenchePA(objProduto)
            If lErro <> SUCESSO Then gError 185270
            
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185271

    Saida_Celula_PAProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_PAProduto:

    Saida_Celula_PAProduto = gErr

    Select Case gErr

        Case 185265, 185266, 185267, 185270, 185271
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 185268
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 185269
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", PAProduto.Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = PAProduto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185272)

    End Select

    Exit Function

End Function

Private Sub BotaoRoteiros_Click()

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoRoteiros_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridPA.Row = 0 Then gError 185273

    lErro = CF("Produto_Formata", GridPA.TextMatrix(GridPA.Row, iGrid_PAProduto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 185274

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
    
    If Len(Trim(GridPA.TextMatrix(GridPA.Row, iGrid_PAVersao_Col))) <> 0 Then
        objRoteirosDeFabricacao.sVersao = GridPA.TextMatrix(GridPA.Row, iGrid_PAVersao_Col)
    End If

    Call Chama_Tela("RoteirosDeFabricacao", objRoteirosDeFabricacao)

    Exit Sub

Erro_BotaoRoteiros_Click:

    Select Case gErr
    
        Case 185273
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 185274

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185275)

    End Select

    Exit Sub

End Sub

Private Sub BotaoKit_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoKit_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridPA.Row = 0 Then gError 185276

    lErro = CF("Produto_Formata", GridPA.TextMatrix(GridPA.Row, iGrid_PAProduto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 185277

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objKit.sProdutoRaiz = sProdutoFormatado
    
    If Len(Trim(GridPA.TextMatrix(GridPA.Row, iGrid_PAVersao_Col))) <> 0 Then
        objKit.sVersao = GridPA.TextMatrix(GridPA.Row, iGrid_PAVersao_Col)
    End If

    Call Chama_Tela("Kit", objKit)

    Exit Sub

Erro_BotaoKit_Click:

    Select Case gErr
    
        Case 185276
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 185277

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185278)

    End Select

    Exit Sub

End Sub

Private Sub PAProduto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PAProduto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPA)
End Sub

Private Sub PAProduto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPA)
End Sub

Private Sub PAProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPA.objControle = PAProduto
    lErro = Grid_Campo_Libera_Foco(objGridPA)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PAQuantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PAQuantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPA)
End Sub

Private Sub PAQuantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPA)
End Sub

Private Sub PAQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPA.objControle = PAQuantidade
    lErro = Grid_Campo_Libera_Foco(objGridPA)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PADescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PADescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPA)
End Sub

Private Sub PADescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPA)
End Sub

Private Sub PADescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPA.objControle = PADescricao
    lErro = Grid_Campo_Libera_Foco(objGridPA)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PAVersao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PAVersao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPA)
End Sub

Private Sub PAVersao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPA)
End Sub

Private Sub PAVersao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPA.objControle = PAVersao
    lErro = Grid_Campo_Libera_Foco(objGridPA)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PAObservacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PAObservacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPA)
End Sub

Private Sub PAObservacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPA)
End Sub

Private Sub PAObservacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPA.objControle = PAObservacao
    lErro = Grid_Campo_Libera_Foco(objGridPA)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPProduto_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPProduto_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMP(iIndex))
End Sub

Private Sub MPProduto_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP(iIndex))
End Sub

Private Sub MPProduto_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMP(iIndex).objControle = MPProduto(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMP(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPQuantidade_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPQuantidade_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMP(iIndex))
End Sub

Private Sub MPQuantidade_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP(iIndex))
End Sub

Private Sub MPQuantidade_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMP(iIndex).objControle = MPQuantidade(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMP(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPDescricao_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPDescricao_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMP(iIndex))
End Sub

Private Sub MPDescricao_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP(iIndex))
End Sub

Private Sub MPDescricao_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMP(iIndex).objControle = MPDescricao(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMP(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPCusto_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPCusto_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMP(iIndex))
End Sub

Private Sub MPCusto_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP(iIndex))
End Sub

Private Sub MPCusto_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMP(iIndex).objControle = MPCusto(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMP(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoMP_Click(iIndex As Integer)

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMP_Click

    If Me.ActiveControl Is MPProduto(iIndex) Then
        sProduto = MPProduto(iIndex).Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridMP(iIndex).Row = 0 Then gError 185279
        
        sProduto = GridMP(iIndex).TextMatrix(GridMP(iIndex).Row, iGrid_MPProduto_Col(iIndex))
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 185280
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    iIndiceInfCalcAtual = iIndex
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutosKitLista", colSelecao, objProduto, objEventoMP)
    
    Exit Sub

Erro_BotaoMP_Click:

    Select Case gErr

        Case 185279
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 185280
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185281)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMP_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim objProdutoKit As New ClassProdutoKit

On Error GoTo Erro_objEventoMP_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 185282
        
    MPProduto(iIndiceInfCalcAtual).PromptInclude = False
    MPProduto(iIndiceInfCalcAtual).Text = sProdutoMascarado
    MPProduto(iIndiceInfCalcAtual).PromptInclude = True
        
    If Not (Me.ActiveControl Is MPProduto(iIndiceInfCalcAtual)) Then
        
        GridMP(iIndiceInfCalcAtual).TextMatrix(GridMP(iIndiceInfCalcAtual).Row, iGrid_MPProduto_Col(iIndiceInfCalcAtual)) = MPProduto(iIndiceInfCalcAtual).Text

        lErro = ProdutoLinha_PreencheMP(objProduto)
        If lErro <> SUCESSO Then gError 185283
        
    End If
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMP_evSelecao:

    Select Case gErr

        Case 185282, 185283
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185284)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_MPProduto(objGridInt As AdmGrid) As Long
'faz a critica da celula de proddduto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim sProdutoMascarado As String

On Error GoTo Erro_Saida_Celula_MPProduto

    Set objGridInt.objControle = MPProduto(iIndiceInfCalcAtual)

    lErro = CF("Produto_Formata", MPProduto(iIndiceInfCalcAtual).Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 185285
    
    'se o produto foi preenchido
    If Len(Trim(MPProduto(iIndiceInfCalcAtual).ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica2", MPProduto(iIndiceInfCalcAtual).Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 185286
        
        'mascara produto escolhido
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 185287

        MPProduto(iIndiceInfCalcAtual).PromptInclude = False
        MPProduto(iIndiceInfCalcAtual).Text = sProdutoMascarado
        MPProduto(iIndiceInfCalcAtual).PromptInclude = True
        
        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
            'se é um produto gerencial  ==> erro
            If lErro = 25043 Then gError 185288
            
            'se não está cadastrado
            If lErro = 25041 Then gError 185289
        
             'Preenche a linha do grid
            lErro = ProdutoLinha_PreencheMP(objProduto)
            If lErro <> SUCESSO Then gError 185290
            
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21985

    Saida_Celula_MPProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_MPProduto:

    Saida_Celula_MPProduto = gErr

    Select Case gErr

        Case 185285, 185286, 185287, 185290
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 185288
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 185289
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", MPProduto(iIndiceInfCalcAtual).Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = MPProduto(iIndiceInfCalcAtual).Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185291)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MPQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MPQuantidade

    Set objGridInt.objControle = MPQuantidade(iIndiceInfCalcAtual)
    
    'se a quantidade foi preenchida
    If Len(MPQuantidade(iIndiceInfCalcAtual).ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MPQuantidade(iIndiceInfCalcAtual).Text)
        If lErro <> SUCESSO Then gError 185292
    
        MPQuantidade(iIndiceInfCalcAtual).Text = Formata_Estoque(MPQuantidade(iIndiceInfCalcAtual).Text)
    
        GridMP(iIndiceInfCalcAtual).TextMatrix(GridMP(iIndiceInfCalcAtual).Row, iGrid_MPCustoT_Col(iIndiceInfCalcAtual)) = Format(StrParaDbl(GridMP(iIndiceInfCalcAtual).TextMatrix(GridMP(iIndiceInfCalcAtual).Row, iGrid_MPCusto_Col(iIndiceInfCalcAtual))) * StrParaDbl(MPQuantidade(iIndiceInfCalcAtual).Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185293
    
    Call Soma_Coluna_Grid(objGridInt, iGrid_MPCustoT_Col(iIndiceInfCalcAtual), MPCustoTotal(iIndiceInfCalcAtual))
    
    Saida_Celula_MPQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MPQuantidade:

    Saida_Celula_MPQuantidade = gErr

    Select Case gErr
    
        Case 185292
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MPQuantidade(iIndiceInfCalcAtual).SetFocus
    
        Case 185293
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185294)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MPCusto(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MPCusto

    Set objGridInt.objControle = MPCusto((iIndiceInfCalcAtual))
    
    'se a quantidade foi preenchida
    If Len(MPCusto((iIndiceInfCalcAtual)).ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MPCusto((iIndiceInfCalcAtual)).Text)
        If lErro <> SUCESSO Then gError 185295
    
        MPCusto((iIndiceInfCalcAtual)).Text = Format(MPCusto((iIndiceInfCalcAtual)).Text, "STANDARD")
    
        GridMP(iIndiceInfCalcAtual).TextMatrix(GridMP(iIndiceInfCalcAtual).Row, iGrid_MPCustoT_Col(iIndiceInfCalcAtual)) = Format(StrParaDbl(GridMP(iIndiceInfCalcAtual).TextMatrix(GridMP(iIndiceInfCalcAtual).Row, iGrid_MPQuantidade_Col(iIndiceInfCalcAtual))) * StrParaDbl(MPCusto((iIndiceInfCalcAtual)).Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185296

    Call Soma_Coluna_Grid(objGridInt, iGrid_MPCustoT_Col(iIndiceInfCalcAtual), MPCustoTotal(iIndiceInfCalcAtual))
    
    Saida_Celula_MPCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_MPCusto:

    Saida_Celula_MPCusto = gErr

    Select Case gErr
           
        Case 185295
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MPCusto(iIndiceInfCalcAtual).SetFocus

        Case 185296
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185297)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MPDescricao(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MPDescricao

    Set objGridInt.objControle = MPDescricao(iIndiceInfCalcAtual)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185298

    Saida_Celula_MPDescricao = SUCESSO

    Exit Function

Erro_Saida_Celula_MPDescricao:

    Saida_Celula_MPDescricao = gErr

    Select Case gErr

        Case 185298
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185299)

    End Select

    Exit Function

End Function

Private Sub MOCodigo_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOCodigo_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMO(iIndex))
End Sub

Private Sub MOCodigo_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO(iIndex))
End Sub

Private Sub MOCodigo_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMO(iIndex).objControle = MOCodigo(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMO(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOQuantidade_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOQuantidade_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMO(iIndex))
End Sub

Private Sub MOQuantidade_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO(iIndex))
End Sub

Private Sub MOQuantidade_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMO(iIndex).objControle = MOQuantidade(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMO(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MODescricao_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MODescricao_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMO(iIndex))
End Sub

Private Sub MODescricao_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO(iIndex))
End Sub

Private Sub MODescricao_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMO(iIndex).objControle = MODescricao(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMO(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOCusto_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOCusto_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMO(iIndex))
End Sub

Private Sub MOCusto_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO(iIndex))
End Sub

Private Sub MOCusto_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMO(iIndex).objControle = MOCusto(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMO(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOHoras_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOHoras_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMO(iIndex))
End Sub

Private Sub MOHoras_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO(iIndex))
End Sub

Private Sub MOHoras_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMO(iIndex).objControle = MOHoras(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMO(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_MOCodigo(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Saida_Celula_MOCodigo

    Set objGridInt.objControle = MOCodigo(iIndiceInfCalcAtual)
    
    'Se o campo foi preenchido
    If Len(MOCodigo(iIndiceInfCalcAtual).Text) > 0 Then
        
        objTiposDeMaodeObra.iCodigo = StrParaInt(MOCodigo(iIndiceInfCalcAtual).Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 185300
    
        If lErro <> SUCESSO Then gError 185301

        GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MODescricao_Col(iIndiceInfCalcAtual)) = objTiposDeMaodeObra.sDescricao
        GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCusto_Col(iIndiceInfCalcAtual)) = Format(objTiposDeMaodeObra.dCustoHora, "STANDARD")
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO((iIndiceInfCalcAtual)).Row - GridMO((iIndiceInfCalcAtual)).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21994

    Saida_Celula_MOCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_MOCodigo:

    Saida_Celula_MOCodigo = gErr

    Select Case gErr
    
        Case 185300
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 185301
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185302)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MOQuantidade

    Set objGridInt.objControle = MOQuantidade(iIndiceInfCalcAtual)
    
    'se a quantidade foi preenchida
    If Len(MOQuantidade(iIndiceInfCalcAtual).ClipText) > 0 Then

        lErro = Valor_Inteiro_Critica(MOQuantidade(iIndiceInfCalcAtual).Text)
        If lErro <> SUCESSO Then gError 185303
        
        GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCustoT_Col(iIndiceInfCalcAtual)) = Format(StrParaDbl(GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCusto_Col(iIndiceInfCalcAtual))) * StrParaDbl(GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOHoras_Col(iIndiceInfCalcAtual))) * StrParaDbl(MOQuantidade(iIndiceInfCalcAtual).Text), "STANDARD")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185304

    Saida_Celula_MOQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MOQuantidade:

    Saida_Celula_MOQuantidade = gErr

    Select Case gErr
    
        Case 185303
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MOQuantidade(iIndiceInfCalcAtual).SetFocus

        Case 185304
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185305)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOHoras(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MOHoras

    Set objGridInt.objControle = MOHoras(iIndiceInfCalcAtual)
    
    'se a quantidade foi preenchida
    If Len(MOHoras(iIndiceInfCalcAtual).ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MOHoras(iIndiceInfCalcAtual).Text)
        If lErro <> SUCESSO Then gError 185306
    
        MOHoras(iIndiceInfCalcAtual).Text = Formata_Estoque(MOHoras(iIndiceInfCalcAtual).Text)
        
        GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCustoT_Col(iIndiceInfCalcAtual)) = Format(StrParaDbl(GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCusto_Col(iIndiceInfCalcAtual))) * StrParaDbl(GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOQuantidade_Col(iIndiceInfCalcAtual))) * StrParaDbl(MOHoras(iIndiceInfCalcAtual).Text), "STANDARD")
    
    End If
    
    Call Soma_Coluna_Grid(objGridInt, iGrid_MOCustoT_Col(iIndiceInfCalcAtual), MOCustoTotal(iIndiceInfCalcAtual))

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185307

    Saida_Celula_MOHoras = SUCESSO

    Exit Function

Erro_Saida_Celula_MOHoras:

    Saida_Celula_MOHoras = gErr

    Select Case gErr
    
        Case 185306
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MOHoras(iIndiceInfCalcAtual).SetFocus
        
        Case 185307
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185308)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOCusto(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MOCusto

    Set objGridInt.objControle = MOCusto((iIndiceInfCalcAtual))
    
    'se a quantidade foi preenchida
    If Len(MOCusto((iIndiceInfCalcAtual)).ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MOCusto((iIndiceInfCalcAtual)).Text)
        If lErro <> SUCESSO Then gError 185309
    
        MOCusto((iIndiceInfCalcAtual)).Text = Format(MOCusto((iIndiceInfCalcAtual)).Text, "STANDARD")
    
        GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCustoT_Col(iIndiceInfCalcAtual)) = Format(StrParaDbl(GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOHoras_Col(iIndiceInfCalcAtual))) * StrParaDbl(GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOQuantidade_Col(iIndiceInfCalcAtual))) * StrParaDbl(MOCusto((iIndiceInfCalcAtual)).Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185310

    Call Soma_Coluna_Grid(objGridInt, iGrid_MOCustoT_Col(iIndiceInfCalcAtual), MOCustoTotal(iIndiceInfCalcAtual))

    Saida_Celula_MOCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_MOCusto:

    Saida_Celula_MOCusto = gErr

    Select Case gErr
    
        Case 185309
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MOCusto(iIndiceInfCalcAtual).SetFocus

        Case 185310
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185311)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MODescricao(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MODescricao

    Set objGridInt.objControle = MODescricao(iIndiceInfCalcAtual)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185312

    Saida_Celula_MODescricao = SUCESSO

    Exit Function

Erro_Saida_Celula_MODescricao:

    Saida_Celula_MODescricao = gErr

    Select Case gErr

        Case 185312
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185313)

    End Select

    Exit Function

End Function

Private Sub MaqCodigo_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqCodigo_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMaq(iIndex))
End Sub

Private Sub MaqCodigo_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq(iIndex))
End Sub

Private Sub MaqCodigo_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq(iIndex).objControle = MaqCodigo(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMaq(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqQuantidade_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqQuantidade_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMaq(iIndex))
End Sub

Private Sub MaqQuantidade_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq(iIndex))
End Sub

Private Sub MaqQuantidade_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq(iIndex).objControle = MaqQuantidade(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMaq(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqDescricao_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqDescricao_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMaq(iIndex))
End Sub

Private Sub MaqDescricao_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq(iIndex))
End Sub

Private Sub MaqDescricao_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq(iIndex).objControle = MaqDescricao(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMaq(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqCusto_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqCusto_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMaq(iIndex))
End Sub

Private Sub MaqCusto_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq(iIndex))
End Sub

Private Sub MaqCusto_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq(iIndex).objControle = MaqCusto(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMaq(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqHoras_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqHoras_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMaq(iIndex))
End Sub

Private Sub MaqHoras_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq(iIndex))
End Sub

Private Sub MaqHoras_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq(iIndex).objControle = MaqHoras(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMaq(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_MaqCodigo(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_Saida_Celula_MaqCodigo

    Set objGridInt.objControle = MaqCodigo(iIndiceInfCalcAtual)
    
    'Se o campo foi preenchido
    If Len(Trim(MaqCodigo(iIndiceInfCalcAtual).Text)) > 0 Then
    
        Set objMaquinas = New ClassMaquinas
    
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", MaqCodigo(iIndiceInfCalcAtual), objMaquinas)
        If lErro <> SUCESSO Then gError 185314
        
        GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqDescricao_Col(iIndiceInfCalcAtual)) = objMaquinas.sDescricao
        GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCusto_Col(iIndiceInfCalcAtual)) = Format(objMaquinas.dCustoHora, "STANDARD")

        'verifica se precisa preencher o grid com uma nova linha
        If GridMaq(iIndiceInfCalcAtual).Row - GridMaq(iIndiceInfCalcAtual).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185315

    Saida_Celula_MaqCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqCodigo:

    Saida_Celula_MaqCodigo = gErr

    Select Case gErr
    
        Case 185314, 185315
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185316)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqQuantidade

    Set objGridInt.objControle = MaqQuantidade(iIndiceInfCalcAtual)
    
    'se a quantidade foi preenchida
    If Len(MaqQuantidade(iIndiceInfCalcAtual).ClipText) > 0 Then

        lErro = Valor_Inteiro_Critica(MaqQuantidade(iIndiceInfCalcAtual).Text)
        If lErro <> SUCESSO Then gError 185317
        
        GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCustoT_Col(iIndiceInfCalcAtual)) = Format(StrParaDbl(GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCusto_Col(iIndiceInfCalcAtual))) * StrParaDbl(GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqHoras_Col(iIndiceInfCalcAtual))) * StrParaDbl(MaqQuantidade(iIndiceInfCalcAtual).Text), "STANDARD")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185318

    Saida_Celula_MaqQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqQuantidade:

    Saida_Celula_MaqQuantidade = gErr

    Select Case gErr
           
        Case 185317
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MaqQuantidade(iIndiceInfCalcAtual).SetFocus

        Case 185318
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185319)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqHoras(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqHoras

    Set objGridInt.objControle = MaqHoras(iIndiceInfCalcAtual)
    
    'se a quantidade foi preenchida
    If Len(MaqHoras(iIndiceInfCalcAtual).ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MaqHoras(iIndiceInfCalcAtual).Text)
        If lErro <> SUCESSO Then gError 185320
    
        MaqHoras(iIndiceInfCalcAtual).Text = Formata_Estoque(MaqHoras(iIndiceInfCalcAtual).Text)
    
        GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCustoT_Col(iIndiceInfCalcAtual)) = Format(StrParaDbl(GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCusto_Col(iIndiceInfCalcAtual))) * StrParaDbl(GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqQuantidade_Col(iIndiceInfCalcAtual))) * StrParaDbl(MaqHoras(iIndiceInfCalcAtual).Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185321

    Call Soma_Coluna_Grid(objGridInt, iGrid_MaqCustoT_Col(iIndiceInfCalcAtual), MaqCustoTotal(iIndiceInfCalcAtual))

    Saida_Celula_MaqHoras = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqHoras:

    Saida_Celula_MaqHoras = gErr

    Select Case gErr

        Case 185320
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MaqHoras(iIndiceInfCalcAtual).SetFocus

        Case 185321
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185322)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqCusto(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqCusto

    Set objGridInt.objControle = MaqCusto((iIndiceInfCalcAtual))
    
    'se a quantidade foi preenchida
    If Len(MaqCusto((iIndiceInfCalcAtual)).ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MaqCusto((iIndiceInfCalcAtual)).Text)
        If lErro <> SUCESSO Then gError 185323
    
        MaqCusto((iIndiceInfCalcAtual)).Text = Format(MaqCusto((iIndiceInfCalcAtual)).Text, "STANDARD")
    
        GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCustoT_Col(iIndiceInfCalcAtual)) = Format(StrParaDbl(GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqHoras_Col(iIndiceInfCalcAtual))) * StrParaDbl(GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqQuantidade_Col(iIndiceInfCalcAtual))) * StrParaDbl(MaqCusto((iIndiceInfCalcAtual)).Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185324
    
    Call Soma_Coluna_Grid(objGridInt, iGrid_MaqCustoT_Col(iIndiceInfCalcAtual), MaqCustoTotal(iIndiceInfCalcAtual))

    Saida_Celula_MaqCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqCusto:

    Saida_Celula_MaqCusto = gErr

    Select Case gErr

        Case 185323
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MaqCusto(iIndiceInfCalcAtual).SetFocus

        Case 185324
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185325)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqDescricao(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqDescricao

    Set objGridInt.objControle = MaqDescricao(iIndiceInfCalcAtual)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 185326

    Saida_Celula_MaqDescricao = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqDescricao:

    Saida_Celula_MaqDescricao = gErr

    Select Case gErr

        Case 185326
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185327)

    End Select

    Exit Function

End Function

Private Function Atualiza_Indice() As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Atualiza_Indice

    Select Case iFrameAtual
    
        Case FRAME_MP
            iIndiceInfCalcAtual = iFrameMP
        Case FRAME_MO
            iIndiceInfCalcAtual = iFrameMO
        Case FRAME_Maq
            iIndiceInfCalcAtual = iFrameMaq
        Case Else
            iIndiceInfCalcAtual = 0
    End Select

    Atualiza_Indice = SUCESSO

    Exit Function

Erro_Atualiza_Indice:

    Atualiza_Indice = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185328)

    End Select

    Exit Function

End Function

Private Function Recalcula_Necessidades() As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim objRoteiroNecess As ClassRoteiroNecessidade
Dim objRoteiroNecessAux As New ClassRoteiroNecessidade
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim colRoteiroNecess As New Collection
Dim objRoteiroMO As ClassRoteiroMaoDeObra
Dim objRoteiroMP As ClassRoteiroMP
Dim objRoteiroMaq As ClassRoteiroMaquina
Dim objRoteiroIMaq As ClassRoteiroInsumosMaq
Dim objRoteiroMOAux As ClassRoteiroMaoDeObra
Dim objRoteiroMPAux As ClassRoteiroMP
Dim objRoteiroMaqAux As ClassRoteiroMaquina
Dim objRoteiroIMaqAux As ClassRoteiroInsumosMaq
Dim bAchou As Boolean
Dim objEtapaMP As ClassPRJEtapaMateriais
Dim objEtapaMO As ClassPRJEtapaMO
Dim objEtapaMaq As ClassPRJEtapaMaquinas
Dim colMP As New Collection
Dim colMO As New Collection
Dim colMaq As New Collection
Dim objProdutos As New ClassProduto
Dim colSaida As Collection
Dim colCampos As Collection
Dim dCusto As Double
Dim dFator As Double

On Error GoTo Erro_Recalcula_Necessidades

    GL_objMDIForm.MousePointer = vbHourglass

    'Monta a coleção de necessidades
    For iIndice = 1 To objGridPA.iLinhasExistentes
    
        Set objRoteiroNecess = New ClassRoteiroNecessidade
        
        lErro = CF("Produto_Formata", GridPA.TextMatrix(iIndice, iGrid_PAProduto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 185455
        
        objRoteiroNecess.dQuantidade = StrParaDbl(GridPA.TextMatrix(iIndice, iGrid_PAQuantidade_Col))
        objRoteiroNecess.iFilialEmpresa = giFilialEmpresa
        objRoteiroNecess.sProdutoRaiz = sProdutoFormatado
        objRoteiroNecess.sUM = GridPA.TextMatrix(iIndice, iGrid_PAUM_Col)
        objRoteiroNecess.sVersao = GridPA.TextMatrix(iIndice, iGrid_PAVersao_Col)

'''=========================================================
''' Alterar essa função:
''' 1) para ler também KIT

        If objRoteiroNecess.dQuantidade > QTDE_ESTOQUE_DELTA Then

            lErro = CF("NecessidadeRoteiro_Le", objRoteiroNecess, True)
            If lErro <> SUCESSO Then gError 185456
            
        End If
        
        colRoteiroNecess.Add objRoteiroNecess

    Next
    
    'Agrupa a necessidades de todos os produtos acabados
    For Each objRoteiroNecess In colRoteiroNecess
    
        For Each objRoteiroMO In objRoteiroNecess.colMaoDeObra
    
            bAchou = False
            For Each objRoteiroMOAux In objRoteiroNecessAux.colMaoDeObra
                If objRoteiroMOAux.iCodMO = objRoteiroMO.iCodMO Then
                    bAchou = True
                    objRoteiroMOAux.dCustoTotal = objRoteiroMOAux.dCustoTotal + objRoteiroMO.dCustoTotal
                    objRoteiroMOAux.dHoras = objRoteiroMOAux.dHoras + objRoteiroMO.dHoras
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objRoteiroMOAux = New ClassRoteiroMaoDeObra
            
                objRoteiroMOAux.dCustoTotal = objRoteiroMO.dCustoTotal
                objRoteiroMOAux.dCustoUnitario = objRoteiroMO.dCustoUnitario
                objRoteiroMOAux.dHoras = objRoteiroMO.dHoras
                objRoteiroMOAux.iCodMO = objRoteiroMO.iCodMO
                objRoteiroMOAux.sUM = objRoteiroMO.sUM
                objRoteiroMOAux.iQuantidade = objRoteiroMO.iQuantidade
                
                Set objRoteiroMOAux.objMaoDeObra = objRoteiroMO.objMaoDeObra
                
                objRoteiroNecessAux.colMaoDeObra.Add objRoteiroMOAux
            End If
    
        Next
        
        For Each objRoteiroMP In objRoteiroNecess.colMP
    
            bAchou = False
            For Each objRoteiroMPAux In objRoteiroNecessAux.colMP
                If objRoteiroMPAux.sProduto = objRoteiroMP.sProduto Then
                    bAchou = True
                            
                    'Realiza a converção para uma mesma UM
                    lErro = CF("UM_Conversao_Trans", objRoteiroMPAux.objProduto.iClasseUM, objRoteiroIMaq.sUM, objRoteiroMP.sUM, dFator)
                    If lErro <> SUCESSO Then gError 189034
                                        
                    objRoteiroMPAux.dCustoTotal = objRoteiroMPAux.dCustoTotal + objRoteiroMP.dCustoTotal
                    objRoteiroMPAux.dQuantidade = objRoteiroMPAux.dQuantidade + objRoteiroMP.dQuantidade * dFator
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objRoteiroMPAux = New ClassRoteiroMP
                Set objProdutos = New ClassProduto
                
                objProdutos.sCodigo = objRoteiroMP.sProduto
                
                lErro = CF("Produto_Le", objProdutos)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 189033
            
                objRoteiroMPAux.dCustoTotal = objRoteiroMP.dCustoTotal
                objRoteiroMPAux.dCustoUnitario = objRoteiroMP.dCustoUnitario
                objRoteiroMPAux.dQuantidade = objRoteiroMP.dQuantidade
                objRoteiroMPAux.sProduto = objRoteiroMP.sProduto
                objRoteiroMPAux.sUM = objRoteiroMP.sUM
                objRoteiroMPAux.sVersao = objRoteiroMP.sVersao
                
                Set objRoteiroMPAux.objProduto = objProdutos
                
                objRoteiroNecessAux.colMP.Add objRoteiroMPAux
            End If
    
        Next
        
        For Each objRoteiroIMaq In objRoteiroNecess.colInsumosMaquina
    
            bAchou = False
            For Each objRoteiroMPAux In objRoteiroNecessAux.colMP
                If objRoteiroMPAux.sProduto = objRoteiroIMaq.sProduto Then
                    bAchou = True
        
                    'Realiza a converção para uma mesma UM
                    lErro = CF("UM_Conversao_Trans", objRoteiroMPAux.objProduto.iClasseUM, objRoteiroIMaq.sUM, objRoteiroMPAux.sUM, dFator)
                    If lErro <> SUCESSO Then gError 189032
                    
                    Set objRoteiroMPAux.objProduto = objProdutos
                    
                    objRoteiroMPAux.dQuantidade = objRoteiroMPAux.dQuantidade + objRoteiroIMaq.dQuantidade * dFator
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objRoteiroMPAux = New ClassRoteiroMP
                
                Set objProdutos = New ClassProduto
                
                objProdutos.sCodigo = objRoteiroIMaq.sProduto
                
                lErro = CF("Produto_Le", objProdutos)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 185457
                    
                objRoteiroMPAux.dCustoTotal = objRoteiroIMaq.dCustoTotal
                objRoteiroMPAux.dCustoUnitario = objRoteiroIMaq.dCustoUnitario
                objRoteiroMPAux.dQuantidade = objRoteiroIMaq.dQuantidade
                objRoteiroMPAux.sProduto = objRoteiroIMaq.sProduto
                objRoteiroMPAux.sUM = objRoteiroIMaq.sUM
                objRoteiroMPAux.sVersao = ""
                
                Set objRoteiroMPAux.objProduto = objProdutos
                
                objRoteiroNecessAux.colMP.Add objRoteiroMPAux
            End If
    
        Next
        
        For Each objRoteiroMaq In objRoteiroNecess.colMaquinas
    
            bAchou = False
            For Each objRoteiroMaqAux In objRoteiroNecessAux.colMaquinas
                If objRoteiroMaqAux.iMaquina = objRoteiroMaq.iMaquina Then
                    bAchou = True
                    If objRoteiroMaqAux.iQuantidade < objRoteiroMaq.iQuantidade Then
                        objRoteiroMaqAux.iQuantidade = objRoteiroMaq.iQuantidade
                    End If
                    objRoteiroMaqAux.dHoras = objRoteiroMaqAux.dHoras + objRoteiroMaq.dHoras
                    objRoteiroMaqAux.dCustoTotal = objRoteiroMaqAux.dCustoUnitario * objRoteiroMaqAux.dHoras
                                     
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objRoteiroMaqAux = New ClassRoteiroMaquina
            
                objRoteiroMaqAux.dCustoTotal = objRoteiroMaq.dCustoTotal
                objRoteiroMaqAux.dCustoUnitario = objRoteiroMaq.dCustoUnitario
                objRoteiroMaqAux.dHoras = objRoteiroMaq.dHoras
                objRoteiroMaqAux.iMaquina = objRoteiroMaq.iMaquina
                objRoteiroMaqAux.iQuantidade = objRoteiroMaq.iQuantidade
                objRoteiroMaqAux.sUM = objRoteiroMaq.sUM
                
                Set objRoteiroMaqAux.objMaquina = objRoteiroMaq.objMaquina
                
                objRoteiroNecessAux.colMaquinas.Add objRoteiroMaqAux
            End If
    
        Next
    
    Next

    For Each objRoteiroMPAux In objRoteiroNecessAux.colMP

        Set objEtapaMP = New ClassPRJEtapaMateriais
    
        lErro = CF("CustoMedioAtual_Le", objRoteiroMPAux.objProduto.sCodigo, dCusto, giFilialEmpresa)
        If lErro <> SUCESSO Then gError 189030
        
        objEtapaMP.dCusto = dCusto * objRoteiroMPAux.dQuantidade
        objEtapaMP.dQuantidade = objRoteiroMPAux.dQuantidade
        objEtapaMP.sDescricao = objRoteiroMPAux.objProduto.sDescricao
        objEtapaMP.sProduto = objRoteiroMPAux.sProduto
        objEtapaMP.sVersao = objRoteiroMPAux.sVersao
        objEtapaMP.sUM = objRoteiroMPAux.sUM
        
        colMP.Add objEtapaMP

    Next
    
    For Each objRoteiroMOAux In objRoteiroNecessAux.colMaoDeObra

        Set objEtapaMO = New ClassPRJEtapaMO
        
        objEtapaMO.dCusto = objRoteiroMOAux.dCustoTotal
        objEtapaMO.iQuantidade = objRoteiroMOAux.iQuantidade
        objEtapaMO.dHoras = objRoteiroMOAux.dHoras
        objEtapaMO.sDescricao = objRoteiroMOAux.objMaoDeObra.sDescricao
        objEtapaMO.iMaoDeObra = objRoteiroMOAux.iCodMO
        
        colMO.Add objEtapaMO

    Next
    
    For Each objRoteiroMaqAux In objRoteiroNecessAux.colMaquinas

        Set objEtapaMaq = New ClassPRJEtapaMaquinas
                
        objEtapaMaq.dCusto = objRoteiroMaqAux.dCustoTotal
        objEtapaMaq.iQuantidade = objRoteiroMaqAux.iQuantidade
        objEtapaMaq.dHoras = objRoteiroMaqAux.dHoras
        objEtapaMaq.sDescricao = objRoteiroMaqAux.objMaquina.sDescricao
        objEtapaMaq.lNumIntDocMaq = objRoteiroMaqAux.objMaquina.lNumIntDoc
        objEtapaMaq.iMaquina = objRoteiroMaqAux.iMaquina
        
        colMaq.Add objEtapaMaq

    Next
    
    Set colSaida = New Collection
    Set colCampos = New Collection
    
    colCampos.Add "sProduto"
    
    Call Ordena_Colecao(colMP, colSaida, colCampos)
    
    lErro = Traz_MP_Tela_Indice(colSaida, INDICE_CALC_PREV)
    If lErro <> SUCESSO Then gError 185459
    
    Set colSaida = New Collection
    Set colCampos = New Collection
    
    colCampos.Add "iMaoDeObra"
    
    Call Ordena_Colecao(colMO, colSaida, colCampos)
    
    lErro = Traz_MO_Tela_Indice(colSaida, INDICE_CALC_PREV)
    If lErro <> SUCESSO Then gError 185460
    
    Set colSaida = New Collection
    Set colCampos = New Collection
    
    colCampos.Add "iMaquina"
    
    Call Ordena_Colecao(colMaq, colSaida, colCampos)
    
    lErro = Traz_Maq_Tela_Indice(colSaida, INDICE_CALC_PREV)
    If lErro <> SUCESSO Then gError 185458
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Recalcula_Necessidades = SUCESSO

    Exit Function

Erro_Recalcula_Necessidades:

    GL_objMDIForm.MousePointer = vbDefault

    Recalcula_Necessidades = gErr

    Select Case gErr
    
        Case 185455 To 185460, 189030 To 189034

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185461)

    End Select

    Exit Function

End Function

Private Sub BotaoMO_Click(iIndex As Integer)

Dim lErro As Long
Dim colSelecao As New Collection
Dim objTiposDeMaodeObras As New ClassTiposDeMaodeObra

On Error GoTo Erro_BotaoMO_Click

    If Me.ActiveControl Is MOCodigo(iIndex) Then
        objTiposDeMaodeObras.iCodigo = StrParaInt(MOCodigo(iIndex))
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridMO(iIndex).Row = 0 Then gError 185462

        objTiposDeMaodeObras.iCodigo = StrParaInt(GridMO(iIndex).TextMatrix(GridMO(iIndex).Row, iGrid_MOCodigo_Col(iIndex)))
        
    End If

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objTiposDeMaodeObras, objEventoMO)

    Exit Sub

Erro_BotaoMO_Click:

    Select Case gErr
        
        Case 185462
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185463)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMO_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer

On Error GoTo Erro_objEventoMO_evSelecao

    Set objTiposDeMaodeObra = obj1
    
    MOCodigo(iIndiceInfCalcAtual).Text = CStr(objTiposDeMaodeObra.iCodigo)
    
    If Not (Me.ActiveControl Is MOCodigo(iIndiceInfCalcAtual)) Then
    
        GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCodigo_Col(iIndiceInfCalcAtual)) = CStr(objTiposDeMaodeObra.iCodigo)
        GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MODescricao_Col(iIndiceInfCalcAtual)) = objTiposDeMaodeObra.sDescricao
        GridMO(iIndiceInfCalcAtual).TextMatrix(GridMO(iIndiceInfCalcAtual).Row, iGrid_MOCusto_Col(iIndiceInfCalcAtual)) = Format(objTiposDeMaodeObra.dCustoHora, "STANDARD")
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO(iIndiceInfCalcAtual).Row - GridMO(iIndiceInfCalcAtual).FixedRows = objGridMO(iIndiceInfCalcAtual).iLinhasExistentes Then
            objGridMO(iIndiceInfCalcAtual).iLinhasExistentes = objGridMO(iIndiceInfCalcAtual).iLinhasExistentes + 1
        End If
        
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMO_evSelecao:

    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185464)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMaq_Click(iIndex As Integer)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMaq_Click

    Set objMaquinas = New ClassMaquinas

    If Me.ActiveControl Is MaqCodigo(iIndex) Then
        objMaquinas.sNomeReduzido = MaqCodigo(iIndex).Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridMaq(iIndex).Row = 0 Then gError 185465
        objMaquinas.sNomeReduzido = GridMaq(iIndex).TextMatrix(GridMaq(iIndex).Row, iGrid_MaqCodigo_Col(iIndex))
    End If
    
    'Le a Máquina no BD a partir do NomeReduzido
    lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
    If lErro <> SUCESSO And lErro <> 103100 Then gError 185466
    
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaq, , "Nome Reduzido")

    Exit Sub

Erro_BotaoMaq_Click:

    Select Case gErr

        Case 185465
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 185466

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185467)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMaq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_objEventoMaq_evSelecao

    Set objMaquinas = obj1

    'Lê o Maquinas
    lErro = CF("TP_Maquina_Le", MaqCodigo(iIndiceInfCalcAtual), objMaquinas)
    If lErro <> SUCESSO Then gError 185468
    
    'Mostra os dados do Maquinas na tela
    MaqCodigo(iIndiceInfCalcAtual).Text = objMaquinas.sNomeReduzido
    
    If Not (Me.ActiveControl Is MaqCodigo(iIndiceInfCalcAtual)) Then
        GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCodigo_Col(iIndiceInfCalcAtual)) = objMaquinas.sNomeReduzido
        GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqDescricao_Col(iIndiceInfCalcAtual)) = objMaquinas.sDescricao
        GridMaq(iIndiceInfCalcAtual).TextMatrix(GridMaq(iIndiceInfCalcAtual).Row, iGrid_MaqCusto_Col(iIndiceInfCalcAtual)) = Format(objMaquinas.dCustoHora, "STANDARD")
    End If
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridMaq(iIndiceInfCalcAtual).Row - GridMaq(iIndiceInfCalcAtual).FixedRows = objGridMaq(iIndiceInfCalcAtual).iLinhasExistentes Then
        objGridMaq(iIndiceInfCalcAtual).iLinhasExistentes = objGridMaq(iIndiceInfCalcAtual).iLinhasExistentes + 1
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoMaq_evSelecao:

    Select Case gErr

        Case 185468
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185469)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Nome(ByVal iTipo As Integer, sTipo As String)

On Error GoTo Erro_Tipo_Nome

    Select Case iTipo
    
        Case INDICE_INF_PREV
            sTipo = "Previsto - Informado"
            
        Case INDICE_CALC_PREV
            sTipo = "Previsto - Calculado"
        
        Case INDICE_INF_REAL
            sTipo = "Real - Informado"
        
        Case INDICE_CALC_REAL
            sTipo = "Real - Calculado"
    
    End Select

    Exit Sub

Erro_Tipo_Nome:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185495)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMaqTrazer_Click(Index As Integer)

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_BotaoMaqTrazer_Click

    lErro = Move_Maq_Memoria(objEtapa, Index + 1)
    If lErro <> SUCESSO Then gError 189300
    
    lErro = Traz_Maq_Tela_Indice(objEtapa.colMaquinas, Index)
    If lErro <> SUCESSO Then gError 189301

    Exit Sub

Erro_BotaoMaqTrazer_Click:

    Select Case gErr
    
        Case 189300, 189301

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189302)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoMOTrazer_Click(Index As Integer)

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_BotaoMOTrazer_Click

    lErro = Move_MO_Memoria(objEtapa, Index + 1)
    If lErro <> SUCESSO Then gError 189303
    
    lErro = Traz_MO_Tela_Indice(objEtapa.colMaoDeObra, Index)
    If lErro <> SUCESSO Then gError 189304

    Exit Sub

Erro_BotaoMOTrazer_Click:

    Select Case gErr
    
        Case 189303, 189304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189305)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoMPTrazer_Click(Index As Integer)

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_BotaoMPTrazer_Click

    lErro = Move_MP_Memoria(objEtapa, Index + 1)
    If lErro <> SUCESSO Then gError 189306
    
    lErro = Traz_MP_Tela_Indice(objEtapa.colMateriaPrima, Index)
    If lErro <> SUCESSO Then gError 189307

    Exit Sub

Erro_BotaoMPTrazer_Click:

    Select Case gErr
    
        Case 189306, 189307

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189308)

    End Select

    Exit Sub
    
End Sub

Function Soma_Coluna_Grid(ByVal objGrid As AdmGrid, ByVal iColuna As Integer, ByVal objControle As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dValor As Double

On Error GoTo Erro_Soma_Coluna_Grid
            
    For iIndice = 1 To objGrid.iLinhasExistentes
        dValor = dValor + StrParaDbl(objGrid.objGrid.TextMatrix(iIndice, iColuna))
    Next
    
    objControle.Caption = Format(dValor, "STANDARD")
        
    Soma_Coluna_Grid = SUCESSO

    Exit Function

Erro_Soma_Coluna_Grid:

    Soma_Coluna_Grid = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189309)

    End Select

    Exit Function

End Function

Private Sub MaqData_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqData_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMaq(iIndex))
End Sub

Private Sub MaqData_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq(iIndex))
End Sub

Private Sub MaqData_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq(iIndex).objControle = MaqData(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMaq(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqOBS_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqOBS_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMaq(iIndex))
End Sub

Private Sub MaqOBS_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq(iIndex))
End Sub

Private Sub MaqOBS_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq(iIndex).objControle = MaqOBS(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMaq(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPData_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPData_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMP(iIndex))
End Sub

Private Sub MPData_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP(iIndex))
End Sub

Private Sub MPData_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMP(iIndex).objControle = MPData(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMP(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPOBS_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPOBS_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMP(iIndex))
End Sub

Private Sub MPOBS_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP(iIndex))
End Sub

Private Sub MPOBS_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMP(iIndex).objControle = MPOBS(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMP(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOData_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOData_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMO(iIndex))
End Sub

Private Sub MOData_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO(iIndex))
End Sub

Private Sub MOData_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMO(iIndex).objControle = MOData(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMO(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOOBS_Change(iIndex As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOOBS_GotFocus(iIndex As Integer)
    Call Grid_Campo_Recebe_Foco(objGridMO(iIndex))
End Sub

Private Sub MOOBS_KeyPress(iIndex As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO(iIndex))
End Sub

Private Sub MOOBS_Validate(iIndex As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridMO(iIndex).objControle = MOOBS(iIndex)
    lErro = Grid_Campo_Libera_Foco(objGridMO(iIndex))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Saida_Celula_Data(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 192142
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 192143

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case 192142 To 192143
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192144)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 192140

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 192140
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192141)

    End Select

    Exit Function

End Function

Private Function Carrega_Usuarios(ByVal objCombo As Object) As Long

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objUsuarios In colUsuarios
        objCombo.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190457)

    End Select

    Exit Function

End Function

Private Sub BotaoCadVistorias_Click()

Dim objVistPRJ As New ClassPRJEtapaVistorias

    objVistPRJ.lNumIntPRJEtapa = glNumIntPRJEtapa
    
    Call Chama_Tela("VistoriaPRJ", objVistPRJ)

End Sub

Private Sub BotaoConVistorias_Click()

Dim objVistPRJ As New ClassPRJEtapaVistorias
Dim colSelecao As New Collection
    
    colSelecao.Add glNumIntPRJEtapa

    Call Chama_Tela("VistoriaPRJLista", colSelecao, objVistPRJ, objEventoCodigo, "NumIntPRJEtapa = ?")

End Sub
