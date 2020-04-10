VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Projetos 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(1)"
      Height          =   4350
      Index           =   1
      Left            =   120
      TabIndex        =   70
      Top             =   675
      Width           =   9255
      Begin VB.Frame Frame2 
         Caption         =   "Outros"
         Height          =   1980
         Index           =   3
         Left            =   45
         TabIndex        =   81
         Top             =   2355
         Width           =   9075
         Begin VB.TextBox Observacao 
            Height          =   330
            Left            =   1725
            MaxLength       =   255
            TabIndex        =   10
            Top             =   1500
            Width           =   7245
         End
         Begin VB.TextBox Responsavel 
            Height          =   330
            Left            =   1725
            TabIndex        =   7
            Top             =   255
            Width           =   2610
         End
         Begin VB.TextBox Objetivo 
            Height          =   330
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   8
            Top             =   675
            Width           =   7245
         End
         Begin VB.TextBox Justificativa 
            Height          =   330
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   9
            Top             =   1080
            Width           =   7245
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   83
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
            TabIndex        =   82
            Top             =   1140
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   630
         Index           =   6
         Left            =   45
         TabIndex        =   78
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
            TabIndex        =   80
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
            TabIndex        =   79
            Top             =   255
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1605
         Index           =   2
         Left            =   45
         TabIndex        =   71
         Top             =   30
         Width           =   9075
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
            TabIndex        =   138
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
         Begin VB.Label Demonstrativo 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   6990
            TabIndex        =   139
            Top             =   150
            Width           =   1890
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            TabIndex        =   73
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
            TabIndex        =   72
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
      Left            =   150
      TabIndex        =   69
      Top             =   660
      Visible         =   0   'False
      Width           =   9255
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   41
         Top             =   1485
         Visible         =   0   'False
         Width           =   3570
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   40
         Top             =   1070
         Visible         =   0   'False
         Width           =   3570
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   42
         Top             =   1900
         Visible         =   0   'False
         Width           =   3570
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   3570
      End
      Begin VB.CommandButton BotaoDadosCustDel 
         Height          =   405
         Left            =   6855
         Picture         =   "ProjetosRL.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3870
         Width           =   435
      End
      Begin VB.ComboBox Controles 
         Height          =   315
         ItemData        =   "ProjetosRL.ctx":04B6
         Left            =   7380
         List            =   "ProjetosRL.ctx":04C6
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   3915
         Width           =   1770
      End
      Begin VB.CommandButton BotaoDadosCustNovo 
         Height          =   405
         Left            =   6390
         Picture         =   "ProjetosRL.ctx":04E6
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3870
         Width           =   435
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Index           =   1
         Left            =   8685
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1485
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
         Left            =   7515
         TabIndex        =   55
         Top             =   1485
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
         Left            =   8685
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1900
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
         Left            =   7515
         TabIndex        =   57
         Top             =   1900
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
         Left            =   8685
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   7515
         TabIndex        =   50
         Top             =   240
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
         Left            =   8685
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   655
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
         Left            =   7515
         TabIndex        =   52
         Top             =   655
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
         Left            =   8685
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   3150
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Index           =   5
         Left            =   7515
         TabIndex        =   61
         Top             =   3145
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
         Left            =   1800
         TabIndex        =   43
         Top             =   2315
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
         Left            =   1800
         TabIndex        =   44
         Top             =   2730
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
         Left            =   1785
         TabIndex        =   45
         Top             =   3145
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
         Left            =   1785
         TabIndex        =   46
         Top             =   3560
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
         Left            =   1785
         TabIndex        =   47
         Top             =   3975
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
         Left            =   7515
         TabIndex        =   59
         Top             =   2315
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   7515
         TabIndex        =   60
         Top             =   2730
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   7515
         TabIndex        =   54
         Top             =   1070
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         TabIndex        =   48
         Top             =   2315
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
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
         TabIndex        =   49
         Top             =   2730
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin VB.TextBox Texto 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   39
         Top             =   655
         Visible         =   0   'False
         Width           =   3570
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
         TabIndex        =   136
         Top             =   2820
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
         TabIndex        =   135
         Top             =   2400
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Cif .:"
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
         Left            =   6075
         TabIndex        =   134
         Top             =   1170
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Peso Liquido.:"
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
         Left            =   6075
         TabIndex        =   133
         Top             =   2850
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Peso Bruto.:"
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
         Left            =   6075
         TabIndex        =   132
         Top             =   2400
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
         Left            =   495
         TabIndex        =   131
         Top             =   4080
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
         Left            =   495
         TabIndex        =   130
         Top             =   3645
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
         Left            =   495
         TabIndex        =   129
         Top             =   3240
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Caixas: "
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
         Left            =   585
         TabIndex        =   128
         Top             =   2820
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantidade:"
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
         Left            =   525
         TabIndex        =   127
         Top             =   2385
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No. DI/DDE/DTA: "
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
         Left            =   120
         TabIndex        =   126
         Top             =   1560
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Container nr.:"
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
         Left            =   540
         TabIndex        =   124
         Top             =   1995
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No. fatura:"
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
         Left            =   780
         TabIndex        =   123
         Top             =   1170
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No. Conhecim.:"
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
         Left            =   390
         TabIndex        =   63
         Top             =   315
         Visible         =   0   'False
         Width           =   1320
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
         Left            =   6225
         TabIndex        =   122
         Top             =   3240
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Chegada:"
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
         Left            =   6060
         TabIndex        =   121
         Top             =   750
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data de Embarque:"
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
         Left            =   5700
         TabIndex        =   120
         Top             =   315
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data de Registro DI:"
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
         Left            =   5580
         TabIndex        =   119
         Top             =   1575
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Desembaraço:"
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
         Left            =   5670
         TabIndex        =   118
         Top             =   1995
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Navio:"
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
         Left            =   1140
         TabIndex        =   125
         Top             =   735
         Visible         =   0   'False
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Escopo"
      Height          =   4290
      Index           =   3
      Left            =   150
      TabIndex        =   111
      Top             =   720
      Visible         =   0   'False
      Width           =   9120
      Begin VB.TextBox EscExclusoes 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   3615
         Width           =   6585
      End
      Begin VB.TextBox EscPremissas 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   2895
         Width           =   6585
      End
      Begin VB.TextBox EscRestricoes 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   2190
         Width           =   6585
      End
      Begin VB.TextBox EscFatores 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   1470
         Width           =   6585
      End
      Begin VB.TextBox EscExpectativa 
         Height          =   645
         Left            =   2400
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   750
         Width           =   6585
      End
      Begin VB.TextBox EscDescricao 
         Height          =   645
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   32
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   112
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
      TabIndex        =   86
      Top             =   660
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame2 
         Caption         =   "Previsão"
         Height          =   2025
         Index           =   0
         Left            =   285
         TabIndex        =   93
         Top             =   45
         Width           =   8730
         Begin VB.Frame Frame2 
            Caption         =   "Calculado"
            Height          =   810
            Index           =   5
            Left            =   135
            TabIndex        =   98
            Top             =   1125
            Width           =   8490
            Begin VB.Label DataFimCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6645
               TabIndex        =   108
               Top             =   315
               Width           =   1170
            End
            Begin VB.Label DataInicioCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4125
               TabIndex        =   107
               Top             =   315
               Width           =   1170
            End
            Begin VB.Label Duracao 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1065
               TabIndex        =   106
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
               TabIndex        =   104
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
               TabIndex        =   103
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
               TabIndex        =   102
               Top             =   345
               Width           =   795
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Informado"
            Height          =   810
            Index           =   4
            Left            =   135
            TabIndex        =   94
            Top             =   240
            Width           =   8490
            Begin MSMask.MaskEdBox Intervalo 
               Height          =   315
               Left            =   1005
               TabIndex        =   22
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
               TabIndex        =   24
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
               TabIndex        =   23
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
               TabIndex        =   26
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
               TabIndex        =   137
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
               TabIndex        =   97
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
               TabIndex        =   96
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
               TabIndex        =   95
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
         TabIndex        =   87
         Top             =   2190
         Width           =   8730
         Begin VB.Frame Frame2 
            Caption         =   "Calculado"
            Height          =   780
            Index           =   8
            Left            =   135
            TabIndex        =   92
            Top             =   1140
            Width           =   8520
            Begin VB.Label DataFimRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   6915
               TabIndex        =   110
               Top             =   285
               Width           =   1170
            End
            Begin VB.Label DataInicioRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4380
               TabIndex        =   109
               Top             =   285
               Width           =   1170
            End
            Begin VB.Label PercCompRealCalc 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2265
               TabIndex        =   105
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
               TabIndex        =   101
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
               TabIndex        =   100
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
               TabIndex        =   99
               Top             =   345
               Width           =   825
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Informado"
            Height          =   780
            Index           =   7
            Left            =   120
            TabIndex        =   88
            Top             =   270
            Width           =   8520
            Begin MSComCtl2.UpDown UpDownDataInicioReal 
               Height          =   300
               Left            =   5550
               TabIndex        =   29
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
               TabIndex        =   28
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
               TabIndex        =   31
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
            Begin MSMask.MaskEdBox PercCompRealInf 
               Height          =   315
               Left            =   2310
               TabIndex        =   27
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
               TabIndex        =   91
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
               TabIndex        =   90
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
               TabIndex        =   89
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
      Height          =   615
      Left            =   8100
      TabIndex        =   17
      ToolTipText     =   "Documentos associados ao Projeto Projeto"
      Top             =   5250
      Width           =   1305
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7245
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   67
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ProjetosRL.ctx":09F8
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ProjetosRL.ctx":0B52
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ProjetosRL.ctx":0CDC
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ProjetosRL.ctx":120E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4755
      Left            =   90
      TabIndex        =   68
      Top             =   345
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   8387
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
      Height          =   615
      Left            =   6765
      TabIndex        =   16
      ToolTipText     =   "Fluxo de Caixa"
      Top             =   5250
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
      Height          =   615
      Left            =   5430
      TabIndex        =   15
      ToolTipText     =   "Contratos"
      Top             =   5250
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
      Height          =   615
      Left            =   1425
      TabIndex        =   12
      ToolTipText     =   "Cronograma do Projeto"
      Top             =   5250
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
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      ToolTipText     =   "Cronograma Físico Financeiro do Projeto"
      Top             =   5250
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
      Height          =   615
      Left            =   4095
      TabIndex        =   14
      ToolTipText     =   "Propostas"
      Top             =   5250
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
      Height          =   615
      Left            =   90
      TabIndex        =   11
      ToolTipText     =   "Organograma do Projeto"
      Top             =   5250
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
    
    'Rocha Log
    Set objCT.gobjInfoUsu = New CTProjetosVGRL
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTProjetosRL

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

