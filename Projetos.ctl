VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Projetos 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(1)"
      Height          =   4680
      Index           =   1
      Left            =   120
      TabIndex        =   72
      Top             =   675
      Width           =   9255
      Begin VB.Frame Frame2 
         Caption         =   "Outros"
         Height          =   2310
         Index           =   3
         Left            =   45
         TabIndex        =   83
         Top             =   2310
         Width           =   9075
         Begin VB.ComboBox Segmento 
            Height          =   315
            Left            =   1710
            TabIndex        =   10
            Top             =   1065
            Width           =   7245
         End
         Begin VB.ComboBox Objetivo 
            Height          =   315
            Left            =   1725
            TabIndex        =   9
            Top             =   660
            Width           =   7245
         End
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
            Left            =   5145
            TabIndex        =   8
            Top             =   300
            Width           =   2400
         End
         Begin VB.ComboBox Responsavel 
            Height          =   315
            Left            =   1725
            TabIndex        =   7
            Top             =   255
            Width           =   3210
         End
         Begin VB.TextBox Observacao 
            Height          =   330
            Left            =   1725
            MaxLength       =   255
            TabIndex        =   12
            Top             =   1920
            Width           =   7245
         End
         Begin VB.TextBox Justificativa 
            Height          =   330
            Left            =   1725
            MaxLength       =   255
            TabIndex        =   11
            Top             =   1500
            Width           =   7230
         End
         Begin VB.Label Label1 
            Caption         =   "Segmento:"
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
            Index           =   24
            Left            =   690
            TabIndex        =   142
            Top             =   1125
            Width           =   900
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
            TabIndex        =   87
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
            Height          =   255
            Index           =   5
            Left            =   510
            TabIndex        =   86
            Top             =   2010
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
            TabIndex        =   85
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
            TabIndex        =   84
            Top             =   1560
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   630
         Index           =   6
         Left            =   45
         TabIndex        =   80
         Top             =   1680
         Width           =   9075
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5445
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
            TabIndex        =   82
            Top             =   255
            Width           =   660
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
            Left            =   4890
            TabIndex        =   81
            Top             =   255
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1605
         Index           =   2
         Left            =   45
         TabIndex        =   73
         Top             =   30
         Width           =   9075
         Begin VB.CommandButton BotaoAnexos 
            Height          =   390
            Left            =   3405
            Picture         =   "Projetos.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   141
            ToolTipText     =   "Anexar Arquivos"
            Top             =   210
            Width           =   420
         End
         Begin VB.CheckBox optEmpToda 
            Caption         =   "Projeto válido para empresa toda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   6960
            TabIndex        =   140
            Top             =   570
            Width           =   2055
         End
         Begin VB.TextBox Descricao 
            Height          =   330
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   4
            Top             =   1110
            Width           =   7245
         End
         Begin MSMask.MaskEdBox NomeReduzido 
            Height          =   315
            Left            =   1710
            TabIndex        =   1
            Top             =   675
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataCriacao 
            Height          =   300
            Left            =   6600
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   660
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataCriacao 
            Height          =   315
            Left            =   5445
            TabIndex        =   2
            Top             =   645
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1710
            TabIndex        =   0
            Top             =   255
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            PromptChar      =   " "
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
            Index           =   0
            Left            =   735
            TabIndex        =   79
            Top             =   1155
            Width           =   930
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   975
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   78
            Top             =   300
            Width           =   660
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
            Left            =   225
            TabIndex        =   77
            Top             =   705
            Width           =   1410
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data de criação:"
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
            Index           =   9
            Left            =   3945
            TabIndex        =   76
            Top             =   705
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Left            =   4785
            TabIndex        =   75
            Top             =   300
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label StatusProjeto 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EM ANDAMENTO"
            Height          =   315
            Left            =   5460
            TabIndex        =   74
            Top             =   255
            Visible         =   0   'False
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(3)"
      Height          =   4350
      Index           =   4
      Left            =   120
      TabIndex        =   71
      Top             =   660
      Visible         =   0   'False
      Width           =   9255
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   4005
         MaxLength       =   255
         TabIndex        =   54
         Top             =   1665
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   4005
         MaxLength       =   255
         TabIndex        =   53
         Top             =   1305
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   4005
         MaxLength       =   255
         TabIndex        =   51
         Top             =   600
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   4005
         MaxLength       =   255
         TabIndex        =   52
         Top             =   945
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4005
         MaxLength       =   255
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.CommandButton BotaoDadosCustDel 
         Height          =   405
         Left            =   6855
         Picture         =   "Projetos.ctx":0196
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3870
         Width           =   435
      End
      Begin VB.ComboBox Controles 
         Height          =   315
         ItemData        =   "Projetos.ctx":064C
         Left            =   7380
         List            =   "Projetos.ctx":065C
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   3915
         Width           =   1770
      End
      Begin VB.CommandButton BotaoDadosCustNovo 
         Height          =   405
         Left            =   6390
         Picture         =   "Projetos.ctx":067C
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3870
         Width           =   435
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   1
         Left            =   2475
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   285
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
         TabIndex        =   40
         Top             =   270
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
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   645
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
         TabIndex        =   42
         Top             =   630
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
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1005
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
         TabIndex        =   44
         Top             =   990
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
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1380
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
         TabIndex        =   46
         Top             =   1365
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1740
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
         TabIndex        =   48
         Top             =   1725
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
         TabIndex        =   56
         Top             =   2190
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
         TabIndex        =   57
         Top             =   2550
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
         TabIndex        =   58
         Top             =   2910
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
         TabIndex        =   59
         Top             =   3270
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
         TabIndex        =   60
         Top             =   3645
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
         TabIndex        =   61
         Top             =   2190
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
         TabIndex        =   62
         Top             =   2550
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
         TabIndex        =   63
         Top             =   2910
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
         TabIndex        =   64
         Top             =   3270
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
         TabIndex        =   65
         Top             =   3645
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
         TabIndex        =   138
         Top             =   3735
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
         TabIndex        =   137
         Top             =   3375
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
         TabIndex        =   136
         Top             =   3015
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
         TabIndex        =   135
         Top             =   2625
         Visible         =   0   'False
         Width           =   1275
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
         TabIndex        =   134
         Top             =   2250
         Visible         =   0   'False
         Width           =   1275
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
         TabIndex        =   133
         Top             =   3690
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
         TabIndex        =   132
         Top             =   3315
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
         TabIndex        =   131
         Top             =   2955
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
         TabIndex        =   130
         Top             =   2595
         Visible         =   0   'False
         Width           =   1170
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
         TabIndex        =   129
         Top             =   2235
         Visible         =   0   'False
         Width           =   1170
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
         TabIndex        =   128
         Top             =   1770
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
         TabIndex        =   127
         Top             =   1410
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
         TabIndex        =   126
         Top             =   1020
         Visible         =   0   'False
         Width           =   660
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
         TabIndex        =   125
         Top             =   690
         Visible         =   0   'False
         Width           =   660
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
         TabIndex        =   55
         Top             =   330
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   124
         Top             =   1770
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
         TabIndex        =   123
         Top             =   1410
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
         TabIndex        =   122
         Top             =   1020
         Visible         =   0   'False
         Width           =   1125
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
         TabIndex        =   121
         Top             =   330
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
         TabIndex        =   120
         Top             =   690
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Escopo"
      Height          =   4290
      Index           =   3
      Left            =   150
      TabIndex        =   113
      Top             =   720
      Visible         =   0   'False
      Width           =   9120
      Begin VB.TextBox EscExclusoes 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   3615
         Width           =   6585
      End
      Begin VB.TextBox EscPremissas 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   2895
         Width           =   6585
      End
      Begin VB.TextBox EscRestricoes 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2190
         Width           =   6585
      End
      Begin VB.TextBox EscFatores 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   1470
         Width           =   6585
      End
      Begin VB.TextBox EscExpectativa 
         Height          =   645
         Left            =   2400
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   750
         Width           =   6585
      End
      Begin VB.TextBox EscDescricao 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   30
         Width           =   6585
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
         TabIndex        =   119
         Top             =   3570
         Width           =   2070
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
         TabIndex        =   118
         Top             =   2865
         Width           =   960
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
         TabIndex        =   117
         Top             =   2175
         Width           =   1155
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
         TabIndex        =   116
         Top             =   1455
         Width           =   1755
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
         TabIndex        =   115
         Top             =   735
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Descrição do Projeto:"
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
         Left            =   435
         TabIndex        =   114
         Top             =   45
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4350
      Index           =   2
      Left            =   120
      TabIndex        =   88
      Top             =   660
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame2 
         Caption         =   "Previsão"
         Height          =   2025
         Index           =   0
         Left            =   285
         TabIndex        =   95
         Top             =   45
         Width           =   8730
         Begin VB.Frame Frame2 
            Caption         =   "Calculado"
            Height          =   810
            Index           =   5
            Left            =   135
            TabIndex        =   100
            Top             =   1125
            Width           =   8490
            Begin VB.Label DataFimCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6645
               TabIndex        =   110
               Top             =   315
               Width           =   1170
            End
            Begin VB.Label DataInicioCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4125
               TabIndex        =   109
               Top             =   315
               Width           =   1170
            End
            Begin VB.Label Duracao 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1065
               TabIndex        =   108
               Top             =   315
               Width           =   1665
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
               TabIndex        =   106
               Top             =   360
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
               Index           =   18
               Left            =   3000
               TabIndex        =   105
               Top             =   360
               Width           =   1035
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
               Index           =   17
               Left            =   195
               TabIndex        =   104
               Top             =   345
               Width           =   795
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Informado"
            Height          =   810
            Index           =   4
            Left            =   135
            TabIndex        =   96
            Top             =   240
            Width           =   8490
            Begin MSMask.MaskEdBox Intervalo 
               Height          =   315
               Left            =   1005
               TabIndex        =   24
               Top             =   315
               Width           =   630
               _ExtentX        =   1111
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
               TabIndex        =   26
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
               TabIndex        =   25
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
               TabIndex        =   28
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
               TabIndex        =   27
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
               Index           =   20
               Left            =   1710
               TabIndex        =   139
               Top             =   375
               Width           =   360
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
               TabIndex        =   99
               Top             =   360
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
               Index           =   2
               Left            =   2985
               TabIndex        =   98
               Top             =   360
               Width           =   1035
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
               TabIndex        =   97
               Top             =   345
               Width           =   795
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Reais"
         Height          =   2025
         Index           =   1
         Left            =   285
         TabIndex        =   89
         Top             =   2190
         Width           =   8730
         Begin VB.Frame Frame2 
            Caption         =   "Calculado"
            Height          =   780
            Index           =   8
            Left            =   135
            TabIndex        =   94
            Top             =   1140
            Width           =   8520
            Begin VB.Label DataFimRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6915
               TabIndex        =   112
               Top             =   285
               Width           =   1170
            End
            Begin VB.Label DataInicioRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4380
               TabIndex        =   111
               Top             =   285
               Width           =   1170
            End
            Begin VB.Label PercCompRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2265
               TabIndex        =   107
               Top             =   300
               Width           =   915
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
               TabIndex        =   103
               Top             =   330
               Width           =   2040
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
               Left            =   3300
               TabIndex        =   102
               Top             =   345
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
               Index           =   14
               Left            =   5940
               TabIndex        =   101
               Top             =   345
               Width           =   825
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Informado"
            Height          =   780
            Index           =   7
            Left            =   120
            TabIndex        =   90
            Top             =   270
            Width           =   8520
            Begin MSComCtl2.UpDown UpDownDataInicioReal 
               Height          =   300
               Left            =   5550
               TabIndex        =   31
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
               TabIndex        =   30
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
               TabIndex        =   33
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
               TabIndex        =   32
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
               TabIndex        =   29
               Top             =   285
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   5
               Format          =   "#0.#0\%"
               PromptChar      =   " "
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
               TabIndex        =   93
               Top             =   330
               Width           =   2040
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
               Left            =   3330
               TabIndex        =   92
               Top             =   345
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
               Index           =   12
               Left            =   5985
               TabIndex        =   91
               Top             =   345
               Width           =   825
            End
         End
      End
   End
   Begin VB.CommandButton BotaoDocRelacs 
      Caption         =   "Documentos Associados"
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
      Left            =   8100
      TabIndex        =   19
      ToolTipText     =   "Documentos associados ao Projeto Projeto"
      Top             =   5415
      Width           =   1305
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7215
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   69
      Top             =   30
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Projetos.ctx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Projetos.ctx":0CE8
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Projetos.ctx":0E72
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Projetos.ctx":13A4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5055
      Left            =   90
      TabIndex        =   70
      Top             =   330
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Escopo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
   Begin VB.CommandButton BotaoFluxo 
      Caption         =   "Fluxo de Caixa"
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
      Left            =   6765
      TabIndex        =   18
      ToolTipText     =   "Fluxo de Caixa"
      Top             =   5415
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton BotaoContrato 
      Caption         =   "Contratos"
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
      Left            =   5430
      TabIndex        =   17
      ToolTipText     =   "Contratos"
      Top             =   5415
      Width           =   1305
   End
   Begin VB.CommandButton BotaoCronograma 
      Caption         =   "Cronograma"
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
      Left            =   1425
      TabIndex        =   14
      ToolTipText     =   "Cronograma do Projeto"
      Top             =   5415
      Width           =   1305
   End
   Begin VB.CommandButton BotaoFisicoFin 
      Caption         =   "Cronograma Fis./Financ."
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
      Left            =   2760
      TabIndex        =   15
      ToolTipText     =   "Cronograma Físico Financeiro do Projeto"
      Top             =   5415
      Width           =   1305
   End
   Begin VB.CommandButton BotaoProposta 
      Caption         =   "Propostas"
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
      Left            =   4095
      TabIndex        =   16
      ToolTipText     =   "Propostas"
      Top             =   5415
      Width           =   1305
   End
   Begin VB.CommandButton BotaoOrganograma 
      Caption         =   "Organograma"
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
      Left            =   90
      TabIndex        =   13
      ToolTipText     =   "Organograma do Projeto"
      Top             =   5415
      Width           =   1305
   End
End
Attribute VB_Name = "Projetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTProjetos
Attribute objCT.VB_VarHelpID = -1

Public Property Set objCTTela(ByVal vData As Object)
    Set objCT = vData
End Property

Public Property Get objCTTela() As Object
    Set objCTTela = objCT
End Property

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoAnexos_Click()
    Call objCT.BotaoAnexos_Click
End Sub

Private Sub BotaoExcluir_Click()
    Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
    Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
    Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
    Call objCT.BotaoLimpar_Click
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
    Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub LabelCliente_Click()
    Call objCT.LabelCliente_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTProjetos
    Set objCT.objUserControl = Me
    
    Call objCT.UserControl_Initialize
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Function Trata_Parametros(Optional objProjetos As ClassProjetos) As Long
     Trata_Parametros = objCT.Trata_Parametros(objProjetos)
End Function

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)
     Call objCT.NomeReduzido_Validate(Cancel)
End Sub

Private Sub NomeReduzido_Change()
     Call objCT.NomeReduzido_Change
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)
     Call objCT.Descricao_Validate(Cancel)
End Sub

Private Sub Descricao_Change()
     Call objCT.Descricao_Change
End Sub

Private Sub UpDownDataCriacao_DownClick()
     Call objCT.UpDownDataCriacao_DownClick
End Sub

Private Sub UpDownDataCriacao_UpClick()
     Call objCT.UpDownDataCriacao_UpClick
End Sub

Private Sub DataCriacao_GotFocus()
     Call objCT.DataCriacao_GotFocus
End Sub

Private Sub DataCriacao_Validate(Cancel As Boolean)
     Call objCT.DataCriacao_Validate(Cancel)
End Sub

Private Sub DataCriacao_Change()
     Call objCT.DataCriacao_Change
End Sub

Private Sub Cliente_GotFocus()
     Call objCT.Cliente_GotFocus
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub Filial_Change()
     Call objCT.Filial_Change
End Sub

Private Sub Responsavel_Validate(Cancel As Boolean)
     Call objCT.Responsavel_Validate(Cancel)
End Sub

Private Sub Responsavel_Change()
     Call objCT.Responsavel_Change
End Sub

Private Sub Objetivo_Validate(Cancel As Boolean)
     Call objCT.Objetivo_Validate(Cancel)
End Sub

Private Sub Objetivo_Change()
     Call objCT.Objetivo_Change
End Sub

Private Sub Justificativa_Validate(Cancel As Boolean)
     Call objCT.Justificativa_Validate(Cancel)
End Sub

Private Sub Justificativa_Change()
     Call objCT.Justificativa_Change
End Sub

Private Sub Observacao_Validate(Cancel As Boolean)
     Call objCT.Observacao_Validate(Cancel)
End Sub

Private Sub Observacao_Change()
     Call objCT.Observacao_Change
End Sub

Private Sub UpDownDataInicio_DownClick()
     Call objCT.UpDownDataInicio_DownClick
End Sub

Private Sub UpDownDataInicio_UpClick()
     Call objCT.UpDownDataInicio_UpClick
End Sub

Private Sub DataInicioInf_GotFocus()
     Call objCT.DataInicioInf_GotFocus
End Sub

Private Sub DataInicioInf_Validate(Cancel As Boolean)
     Call objCT.DataInicioInf_Validate(Cancel)
End Sub

Private Sub DataInicioInf_Change()
     Call objCT.DataInicioInf_Change
End Sub

Private Sub UpDownDataFim_DownClick()
     Call objCT.UpDownDataFim_DownClick
End Sub

Private Sub UpDownDataFim_UpClick()
     Call objCT.UpDownDataFim_UpClick
End Sub

Private Sub DataFimInf_GotFocus()
     Call objCT.DataFimInf_GotFocus
End Sub

Private Sub DataFimInf_Validate(Cancel As Boolean)
     Call objCT.DataFimInf_Validate(Cancel)
End Sub

Private Sub DataFimInf_Change()
     Call objCT.DataFimInf_Change
End Sub

Private Sub UpDownDataInicioReal_DownClick()
     Call objCT.UpDownDataInicioReal_DownClick
End Sub

Private Sub UpDownDataInicioReal_UpClick()
     Call objCT.UpDownDataInicioReal_UpClick
End Sub

Private Sub DataInicioRealInf_GotFocus()
     Call objCT.DataInicioRealInf_GotFocus
End Sub

Private Sub DataInicioRealInf_Validate(Cancel As Boolean)
     Call objCT.DataInicioRealInf_Validate(Cancel)
End Sub

Private Sub DataInicioRealInf_Change()
     Call objCT.DataInicioRealInf_Change
End Sub

Private Sub UpDownDataFimReal_DownClick()
     Call objCT.UpDownDataFimReal_DownClick
End Sub

Private Sub UpDownDataFimReal_UpClick()
     Call objCT.UpDownDataFimReal_UpClick
End Sub

Private Sub DataFimRealInf_GotFocus()
     Call objCT.DataFimRealInf_GotFocus
End Sub

Private Sub DataFimRealInf_Validate(Cancel As Boolean)
     Call objCT.DataFimRealInf_Validate(Cancel)
End Sub

Private Sub DataFimRealInf_Change()
     Call objCT.DataFimRealInf_Change
End Sub

Private Sub PercCompRealInf_Validate(Cancel As Boolean)
     Call objCT.PercCompRealInf_Validate(Cancel)
End Sub

Private Sub PercCompRealInf_GotFocus()
     Call objCT.PercCompRealInf_GotFocus
End Sub

Private Sub PercCompRealInf_Change()
     Call objCT.PercCompRealInf_Change
End Sub

Private Sub LabelCodigo_Click()
     Call objCT.LabelCodigo_Click
End Sub

Private Sub LabelNomeReduzido_Click()
     Call objCT.LabelNomeReduzido_Click
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

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
Private Sub Data_Change(Index As Integer)
     Call objCT.Data_Change(Index)
End Sub

Private Sub Data_GotFocus(Index As Integer)
     Call objCT.Data_GotFocus(Index)
End Sub

Private Sub Data_Validate(Index As Integer, Cancel As Boolean)
     Call objCT.Data_Validate(Index, Cancel)
End Sub

Private Sub Numero_Change(Index As Integer)
     Call objCT.Numero_Change(Index)
End Sub

Private Sub Numero_GotFocus(Index As Integer)
     Call objCT.Numero_GotFocus(Index)
End Sub

Private Sub Numero_Validate(Index As Integer, Cancel As Boolean)
     Call objCT.Numero_Validate(Index, Cancel)
End Sub

Private Sub Valor_Change(Index As Integer)
     Call objCT.Valor_Change(Index)
End Sub

Private Sub Valor_GotFocus(Index As Integer)
     Call objCT.Valor_GotFocus(Index)
End Sub

Private Sub Valor_Validate(Index As Integer, Cancel As Boolean)
     Call objCT.Valor_Validate(Index, Cancel)
End Sub

Private Sub Texto_Change(Index As Integer)
     Call objCT.Texto_Change(Index)
End Sub

Private Sub UpDownData_DownClick(Index As Integer)
     Call objCT.UpDownData_DownClick(Index)
End Sub

Private Sub UpDownData_UpClick(Index As Integer)
     Call objCT.UpDownData_UpClick(Index)
End Sub

Private Sub BotaoDadosCustNovo_Click()
     Call objCT.BotaoDadosCustNovo_Click
End Sub

Private Sub BotaoDadosCustDel_Click()
     Call objCT.BotaoDadosCustDel_Click
End Sub

Private Sub EscDescricao_Change()
     Call objCT.EscDescricao_Change
End Sub

Private Sub EscExclusoes_Change()
     Call objCT.EscExclusoes_Change
End Sub

Private Sub EscExpectativa_Change()
     Call objCT.EscExpectativa_Change
End Sub

Private Sub EscFatores_Change()
     Call objCT.EscFatores_Change
End Sub

Private Sub EscPremissas_Change()
     Call objCT.EscPremissas_Change
End Sub

Private Sub EscRestricoes_Change()
     Call objCT.EscRestricoes_Change
End Sub

Private Sub Intervalo_Change()
     Call objCT.Intervalo_Change
End Sub

Private Sub Intervalo_Validate(Cancel As Boolean)
     Call objCT.Intervalo_Validate(Cancel)
End Sub

Private Sub BotaoOrganograma_Click()
     Call objCT.BotaoOrganograma_Click
End Sub

Private Sub BotaoDocRelacs_Click()
     Call objCT.BotaoDocRelacs_Click
End Sub

Private Sub BotaoCronograma_Click()
     Call objCT.BotaoCronograma_Click
End Sub

Private Sub BotaoFisicoFin_Click()
     Call objCT.BotaoFisicoFin_Click
End Sub

Private Sub BotaoProposta_Click()
     Call objCT.BotaoProposta_Click
End Sub

Private Sub BotaoContrato_Click()
     Call objCT.BotaoContrato_Click
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub Segmento_Validate(Cancel As Boolean)
     Call objCT.Segmento_Validate(Cancel)
End Sub

Private Sub Segmento_Change()
     Call objCT.Segmento_Change
End Sub
